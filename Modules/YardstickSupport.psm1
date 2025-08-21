# Useful Functions

# Write-Log
# Prints both to the console and also to the log
# Param: [String] $Content, [Switch] $Init (Erase log and print big timestamp for new script run)
function Write-Log {
    <#
    .SYNOPSIS
    Writes timestamped log entries to both console and log file.
    
    .DESCRIPTION
    This function writes log messages with timestamps to both the console output
    and a log file. It can also initialize the log file with a header.
    
    .PARAMETER Content
    The content to write to the log.
    
    .PARAMETER Init
    Switch to initialize/clear the log file and write a header.
    #>
    param(
        [String]$Content,
        [Switch]$Init
    )
    
    if (-not $LOG_LOCATION -or -not $LOG_FILE) {
        Write-Warning "LOG_LOCATION or LOG_FILE variables are not set. Cannot write to log."
        if ($Content) {
            Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss") - $Content"
        }
        return
    }
    
    Push-Location $LOG_LOCATION
    try {
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
    }
    catch {
        Write-Warning "Failed to write to log file: $_"
        if ($Content) {
            Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss") - $Content"
        }
    }
    finally {
        Pop-Location
    }
}


# ArrayToString
# Converts an array to a string representation
function ArrayToString {
    <#
    .SYNOPSIS
    Converts an array to a PowerShell array string representation.
    
    .DESCRIPTION
    This function takes an array and converts it to a string representation
    that looks like a PowerShell array literal (@(value1,value2,value3)).
    
    .PARAMETER Array
    The array to convert to string format.
    
    .OUTPUTS
    String representation of the array.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [Array] $Array
    )
    
    if (-not $Array -or $Array.Count -eq 0) {
        return "@()"
    }
    
    $arrayString = "@("
    foreach ($value in $Array) {
        $arrayString = "$($arrayString)$([String]$value),"
    }
    $arrayString = $arrayString.TrimEnd(",")
    $arrayString = "$arrayString)"
    return [String]$arrayString
}

# Get-RedirectedUrl
# Follows redirects and returns the actual URL of a resource
function Get-RedirectedUrl {
    <#
    .SYNOPSIS
    Follows HTTP redirects and returns the final URL.
    
    .DESCRIPTION
    This function follows HTTP redirects to determine the final destination URL
    of a given URL. Useful for handling shortened URLs or redirects.
    
    .PARAMETER URL
    The URL to follow redirects for.
    
    .OUTPUTS
    The final redirected URL as a string.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [String]$URL
    )
    
    $userAgent = [Microsoft.PowerShell.Commands.PSUserAgent]::Chrome
    $httpClient = $null
    
    try {
        $httpClient = [System.Net.Http.HttpClient]::new()
        $httpClient.DefaultRequestHeaders.UserAgent.ParseAdd($userAgent)

        # Get the redirected url object
        $Response = $httpClient.GetAsync($URL).GetAwaiter().GetResult()
        if ($Response.StatusCode -eq "OK") {
            $RedirectedURL = $Response.RequestMessage.RequestUri.AbsoluteUri
        }
        else {
            throw "HTTP request failed with status: $($Response.StatusCode)"
        }
        return $RedirectedURL
    }
    catch {
        Write-Error "Error getting redirected URL for '$URL': $_"
        throw
    }
    finally {
        if ($httpClient) {
            $httpClient.Dispose()
        }
    }
}


# Get-MsiProductCode
# Returns the product code of an MSI file
function Get-MsiProductCode {
    <#
    .SYNOPSIS
    Extracts the ProductCode from an MSI file.
    
    .DESCRIPTION
    This function uses the Windows Installer COM object to read the ProductCode
    property from an MSI database file.
    
    .PARAMETER FilePath
    The full path to the MSI file.
    
    .OUTPUTS
    The MSI ProductCode as a string.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [String]$FilePath
    )
    
    if (-not (Test-Path $FilePath)) {
        throw "MSI file not found: $FilePath"
    }
    
    $windowsInstallerObject = $null
    $msiDatabase = $null
    $view = $null
    
    try {
        # Read property from MSI database
        $windowsInstallerObject = New-Object -ComObject WindowsInstaller.Installer
        $msiDatabase = $windowsInstallerObject.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $windowsInstallerObject, @($FilePath, 0))
        $query = "SELECT Value FROM Property WHERE Property = 'ProductCode'"
        $view = $msiDatabase.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $msiDatabase, ($query))
        $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)
        $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
        
        if ($record) {
            $value = $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1)
        }
        else {
            throw "ProductCode not found in MSI file"
        }
        
        return [String]$value
    }
    catch {
        Write-Error "Error reading ProductCode from MSI file '$FilePath': $_"
        throw
    }
    finally {
        # Clean up COM objects
        if ($view) { 
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($view) 
        }
        if ($msiDatabase) { 
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($msiDatabase) 
        }
        if ($windowsInstallerObject) { 
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($windowsInstallerObject) 
        }
    }
}

# Connect-AutoMSIntuneGraph
# Automatically connects to the Microsoft Graph API using the Intune module
function Connect-AutoMSIntuneGraph {
    <#
    .SYNOPSIS
    Automatically manages Microsoft Intune Graph API connection with token refresh.
    
    .DESCRIPTION
    This function manages the connection to Microsoft Intune Graph API, automatically
    refreshing tokens when they expire or are close to expiring. It uses global
    variables for tenant configuration.
    #>
    
    if (-not $Global:TENANT_ID -or -not $Global:CLIENT_ID -or -not $Global:CLIENT_SECRET) {
        throw "Required global variables not set: TENANT_ID, CLIENT_ID, CLIENT_SECRET"
    }
    
    # Check if the current token is invalid
    if (-not $Global:Token) {
        Write-Log "Getting an Intune Graph Client API token..."
        try {
            $Global:Token = Connect-MSIntuneGraph -TenantID $Global:TENANT_ID -ClientID $Global:CLIENT_ID -ClientSecret $Global:CLIENT_SECRET
        }
        catch {
            Write-Error "Failed to get initial token: $_"
            throw
        }
    }
    elseif ($Global:Token.ExpiresOn.ToLocalTime() -lt (Get-Date)) {
        # If not, get a new token
        Write-Log "Token is expired. Refreshing token..."
        try {
            $Global:Token = Connect-MSIntuneGraph -TenantID $Global:TENANT_ID -ClientID $Global:CLIENT_ID -ClientSecret $Global:CLIENT_SECRET
            Write-Log "Token refreshed. New Token Expires at: $($Global:Token.ExpiresOn.ToLocalTime())"
        }
        catch {
            Write-Error "Failed to refresh expired token: $_"
            throw
        }
    }
    elseif ($Global:Token.ExpiresOn.AddMinutes(-30).ToLocalTime() -lt (Get-Date)) {
        # For whatever reason, this API stops working 10 minutes before a token refresh
        # Set at 30 minutes in case we are uploading large files. 
        Write-Log "Token expires soon - Refreshing token..."
        try {
            # Required to force a refresh
            Clear-MsalTokenCache
            $Global:Token = Connect-MSIntuneGraph -TenantID $Global:TENANT_ID -ClientID $Global:CLIENT_ID -ClientSecret $Global:CLIENT_SECRET
            Write-Log "Token refreshed. New Token Expires at: $($Global:Token.ExpiresOn.ToLocalTime())"
        }
        catch {
            Write-Error "Failed to refresh token: $_"
            throw
        }
    }
    else {
        Write-Log "Token is still valid. Skipping token refresh."
    } 
}


# Move-AssignmentsAndDependencies
# Moves all assignments and dependencies from one application to another
function Move-AssignmentsAndDependencies {
    <#
    .SYNOPSIS
    Moves assignments and dependencies from one Intune application to another.
    
    .DESCRIPTION
    This function transfers all group assignments and application dependencies
    from a source application to a target application, with options to offset
    availability and deadline dates.
    
    .PARAMETER From
    The source application object to move assignments and dependencies from.
    
    .PARAMETER To
    The target application object to move assignments and dependencies to.
    
    .PARAMETER DeadlineDateOffset
    Number of days to offset the deadline date by (default: 0).
    
    .PARAMETER AvailableDateOffset
    Number of days to offset the available date by (default: 0).
    #>
    param(
        [Parameter(Mandatory, Position=0)]
        [System.Object] $From,
        [Parameter(Mandatory, Position=1)]
        [System.Object] $To,
        [Parameter(Position=2)]
        [Int] $DeadlineDateOffset = 0,
        [Parameter(Position=3)]
        [Int] $AvailableDateOffset = 0
    )
    Write-Log "Moving assignments and dependencies from $($From.id) to $($To.id)"
    $FromAssignments = Get-IntuneWin32AppAssignment -Id $From.id
    $FromDependencies = Get-IntuneWin32AppDependency -Id $From.id
    $AvailableDate = (Get-Date).AddDays($AvailableDateOffset).ToString("MM/dd/yyyy")
    $DeadlineDate = (Get-Date).AddDays($DeadlineDateOffset).ToString("MM/dd/yyyy")
    if ($FromAssignments) {
        foreach ($Assignment in $FromAssignments) {
            $maxRetries = 3
            $try = 0
            $successfullyAdded = $false
            while (!$successfullyAdded -and ($try++ -lt $maxRetries)) {
                if ($Assignment.InstallTimeSettings) {
                    $useLocalTime = [bool]$Assignment.InstallTimeSettings.useLocalTime
                    $startDateTime = $Assignment.InstallTimeSettings.startDateTime
                    $deadlineDateTime = $Assignment.InstallTimeSettings.deadlineDateTime
                    if ($null -ne $startDateTime) {
                        Write-Log "StartDateTime: $startDateTime"
                        $startDateTime = Get-Date -Date "$AvailableDate $($startDateTime.ToString("HH:mm"))"
                    }
                    if ($null -ne $deadlineDateTime) {
                        Write-Log "DeadlineDateTime: $deadlineDateTime"
                        $deadlineDateTime = Get-Date -Date "$DeadlineDate $($deadlineDateTime.ToString("HH:mm"))"
                    }
                }
                if ($Assignment.FilterType -eq "none") {
                    if ($Assignment.InstallTimeSettings) {
                        if ($startDateTime -and $deadlineDateTime) {
                            Write-Log "Adding assignment to $($To.id) for group $($Assignment.GroupID) with deadline and available time settings."
                            try {
                                Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $Assignment.GroupID -Intent $Assignment.Intent -Notification $Assignment.Notifications -AvailableTime $startDateTime -DeadlineTime $deadlineDateTime -UseLocalTime $useLocalTime | Out-Null
                                $successfullyAdded = $true
                            }
                            catch {
                                Write-Log "Failed to add assignment to $($To.id) for group $($Assignment.GroupID): $_"
                            }
                        }
                        elseif ($startDateTime) {
                            Write-Log "Adding assignment to $($To.id) for group $($Assignment.GroupID) with available time settings."
                            try {
                                Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $Assignment.GroupID -Intent $Assignment.Intent -Notification $Assignment.Notifications -AvailableTime $startDateTime -UseLocalTime $useLocalTime | Out-Null
                                $successfullyAdded = $true
                            }
                            catch {
                                Write-Log "Failed to add assignment to $($To.id) for group $($Assignment.GroupID): $_"
                            }
                        }
                        elseif ($deadlineDateTime) {
                            Write-Log "Adding assignment to $($To.id) for group $($Assignment.GroupID) with deadline time settings."
                            try {
                                Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $Assignment.GroupID -Intent $Assignment.Intent -Notification $Assignment.Notifications -DeadlineTime $deadlineDateTime -UseLocalTime $useLocalTime | Out-Null
                                $successfullyAdded = $true
                            }
                            catch {
                                Write-Log "Failed to add assignment to $($To.id) for group $($Assignment.GroupID): $_"
                            }
                        }
                        else {
                            Write-Log "Adding assignment to $($To.id) for group $($Assignment.GroupID) without time settings."
                            try {
                                Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $Assignment.GroupID -Intent $Assignment.Intent -Notification $Assignment.Notifications | Out-Null
                                $successfullyAdded = $true
                            }
                            catch {
                                Write-Log "Failed to add assignment to $($To.id) for group $($Assignment.GroupID): $_"
                            }
                        }
                    }
                }
                else {
                    # If there is a filter
                    if ($startDateTime -and $deadlineDateTime) {
                        Write-Log "Adding assignment to $($To.id) for group $($Assignment.GroupID) with deadline and available time settings."
                        try {
                            Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $Assignment.GroupID -Intent $Assignment.Intent -Notification $Assignment.Notifications -AvailableTime $startDateTime -DeadlineTime $deadlineDateTime -UseLocalTime $useLocalTime -FilterMode $Assignment.FilterType -FilterID $Assignment.FilterID | Out-Null
                            $successfullyAdded = $true
                        }
                        catch {
                                Write-Log "Failed to add assignment to $($To.id) for group $($Assignment.GroupID): $_"
                        }
                    }
                    elseif ($startDateTime) {
                        Write-Log "Adding assignment to $($To.id) for group $($Assignment.GroupID) with available time settings."
                        try {
                            Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $Assignment.GroupID -Intent $Assignment.Intent -Notification $Assignment.Notifications -AvailableTime $startDateTime -UseLocalTime $useLocalTime -FilterMode $Assignment.FilterType -FilterID $Assignment.FilterID | Out-Null
                            $successfullyAdded = $true
                        }
                        catch {
                                Write-Log "Failed to add assignment to $($To.id) for group $($Assignment.GroupID): $_"
                        }
                    }
                    elseif ($deadlineDateTime) {
                        Write-Log "Adding assignment to $($To.id) for group $($Assignment.GroupID) with deadline time settings."
                        try {
                            Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $Assignment.GroupID -Intent $Assignment.Intent -Notification $Assignment.Notifications -DeadlineTime $deadlineDateTime -UseLocalTime $useLocalTime -FilterMode $Assignment.FilterType -FilterID $Assignment.FilterID | Out-Null
                            $successfullyAdded = $true
                        }
                        catch {
                                Write-Log "Failed to add assignment to $($To.id) for group $($Assignment.GroupID): $_"
                        }
                    }
                    else {
                        Write-Log "Adding assignment to $($To.id) for group $($Assignment.GroupID) without time settings."
                        try {
                                Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $Assignment.GroupID -Intent $Assignment.Intent -Notification $Assignment.Notifications -FilterMode $Assignment.FilterType -FilterID $Assignment.FilterID | Out-Null
                                $successfullyAdded = $true
                        }
                        catch {
                                Write-Log "Failed to add assignment to $($To.id) for group $($Assignment.GroupID): $_"
                        }
                    }
                } 
            }
            # Remove the old assignment
            if ($successfullyAdded) {
                for ($i = 0; $i -le 3; $i++) {
                    try {
                        Remove-IntuneWin32AppAssignmentGroup -ID $From.id -GroupID $Assignment.GroupID | Out-Null
                        if ($?) {
                            Write-Log "Successfully removed assignment from $($From.id) for group $($Assignment.GroupID)"
                            break
                        }
                    }
                    catch {
                        Write-Log "Failed to remove assignment $($Assignment.GroupID) from group $($From.id):"
                        if ($i -eq 3) {
                            Write-Log "Failed to remove assignment after 3 attempts. Skipping removal."
                            break
                        }
                        else {
                            Write-Log "Retrying removal of assignment from $($From.id) for group $($Assignment.GroupID). Attempt $($i + 1) of 3."
                            Start-Sleep -Seconds 2
                            continue
                        }
                    }
                }
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
function Get-SameAppAllVersions {
    <#
    .SYNOPSIS
    Retrieves all versions of an application sorted from newest to oldest.
    
    .DESCRIPTION
    This function finds all applications with the same display name (including
    versioned names with N- prefix) and returns them sorted by version in
    descending order. Accounts for edge cases where application names might
    be similar to others.
    
    .PARAMETER DisplayName
    The display name of the application to search for.
    
    .OUTPUTS
    Array of application objects sorted by version (newest first).
    #>
    param(
        [Parameter(Mandatory=$true)]
        [String]$DisplayName
    )
    
    # Attempt to connect to Intune up to 3 times
    for ($i = 0; $i -lt 3; $i++) {
        try {
            $AllSimilarApps = Get-IntuneWin32App -DisplayName "$DisplayName" -ErrorAction SilentlyContinue
            break
        }
        catch {
            Write-Log "Error retrieving applications with the name $DisplayName"
            if ($i -eq 2) {
                Write-Log "Intune API failed to retrieve applications after 3 attempts. Exiting."
                exit 1001
            }
            else {
                Write-Log "Retrying to retrieve applications with the name $DisplayName. Attempt $($i + 1) of 3."
                Start-Sleep -Seconds 5
            }
        }
    }
    
    if (-not $AllSimilarApps) {
        Write-Log "No applications found with the name $DisplayName"
        return @()
    }
    
    # Only return apps that have the same name (not ones that look the same) and apps that are pending approval
    $sortable = ($AllSimilarApps | Where-Object {($_.DisplayName -eq $DisplayName) -or ($_.DisplayName -like "$DisplayName (N-*")})
    
    # Pad out each version section to 8 digits of zeroes before sorting, remove any letters, and mash it all together so that versions sort correctly
    return $sortable | Sort-Object {$($($($_.displayVersion -replace "[A-Za-z]", "0") -replace "[\-\+]", ".").split(".") | ForEach-Object {'{0:d8}' -f [int]$_}) -join ''} -Descending
}


# Format-FileDetectionVersion
# Converts a version number to one padded with the appropriate number of zeroes for detection by Intune
function Format-FileDetectionVersion {
    <#
    .SYNOPSIS
    Formats a version number for Intune file detection.
    
    .DESCRIPTION
    Converts a version number to a 4-part version string padded with zeroes
    as required by Intune for file detection rules.
    
    .PARAMETER Version
    The version string to format.
    
    .OUTPUTS
    A properly formatted 4-part version string (e.g., "1.2.3.0").
    #>
    param(
        [Parameter(Mandatory=$true)]
        [String]$Version
    )
    
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
# Returns whether a version is locked based on version lock pattern
function Get-VersionLocked {
    <#
    .SYNOPSIS
    Determines if a version is locked based on a version lock pattern.
    
    .DESCRIPTION
    Checks if a version should be blocked from updating based on a version
    lock pattern. The pattern can use 'x' as wildcards (e.g., "1.2.x" locks
    to version 1.2 but allows any patch version).
    
    .PARAMETER Version
    The version to check against the lock pattern.
    
    .PARAMETER VersionLock
    The version lock pattern. Use 'x' for wildcards.
    
    .OUTPUTS
    $true if the version is locked (should not update), $false otherwise.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [String]$Version,
        [Parameter(Mandatory=$false)]
        [String]$VersionLock
    )
    
    # If versionLock is null or empty, return false (not locked)
    if (-not $VersionLock) {
        return $false
    }
    
    # Compare the version and versionLock
    $versionPattern = $VersionLock -replace "[Xx]{1,}", "[0-9]{1,}"
    $versionPattern = $versionPattern -replace "\.", "\."
    return $Version -notmatch "^$versionPattern"
}


# Compare-AppVersions
# Compares two version strings and returns comparison result
function Compare-AppVersions {
    <#
    .SYNOPSIS
    Compares two application version strings.
    
    .DESCRIPTION
    Compares two version strings numerically and returns -1, 0, or 1
    based on whether the first version is less than, equal to, or
    greater than the second version.
    
    .PARAMETER Version1
    The first version string to compare.
    
    .PARAMETER Version2
    The second version string to compare.
    
    .OUTPUTS
    -1 if Version1 < Version2
     0 if Version1 = Version2
     1 if Version1 > Version2
    #>
    param (
        [Parameter(Mandatory=$true)]
        [String]$Version1,
        [Parameter(Mandatory=$true)]
        [String]$Version2
    )
    
    # Clean version strings to contain only numbers and dots
    $Version1 = $Version1 -replace "[^0-9.]", ""
    $Version2 = $Version2 -replace "[^0-9.]", ""
    
    $version1Components = $Version1.split(".")
    $version2Components = $Version2.split(".")
    $maxLength = [Math]::Max($version1Components.Count, $version2Components.Count)
    
    for ($i = 0; $i -lt $maxLength; $i++) {
        $v1 = if ($i -lt $version1Components.Count) { [int]$version1Components[$i] } else { 0 }
        $v2 = if ($i -lt $version2Components.Count) { [int]$version2Components[$i] } else { 0 }
        
        if ($v1 -lt $v2) {
            return -1
        }
        elseif ($v1 -gt $v2) {
            return 1
        }
    }
    return 0
}

