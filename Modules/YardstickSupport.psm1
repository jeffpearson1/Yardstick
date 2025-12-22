using module .\VersionPro.psm1

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
    
    if (-not $LogLocation -or -not $LogFile) {
        Write-Warning "LOG_LOCATION or LOG_FILE variables are not set. Cannot write to log."
        if ($Content) {
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss") - $Content"
        }
        return
    }
    
    Push-Location $LogLocation
    try {
        if ($Init) {
            if (Test-Path $LogLocation\$LogFile) {
                Remove-Item $LogLocation\$LogFile -Force
            }
            Write-Output "#######################################################" | Out-File $LogFile -Append
            Write-Output "LOGGING STARTED AT $(Get-Date -Format "MM/dd/yyyy HH:mm:ss")" | Out-File $LogFile -Append
            Write-Output "#######################################################" | Out-File $LogFile -Append
        }
        if ($Content) {
            $Content = "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss") - $Content"
            Write-Output $Content | Out-File $LogFile -Append
            Write-Host $Content
        }
    } catch {
        Write-Warning "Failed to write to log file: $_"
        if ($Content) {
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss") - $Content"
        }
    } finally {
        Pop-Location -ErrorAction SilentlyContinue
    }
}



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
        } else {
            throw "HTTP request failed with status: $($Response.StatusCode)"
        }
        return $RedirectedURL
    } catch {
        Write-Error "Error getting redirected URL for '$URL': $_"
        throw
    } finally {
        if ($httpClient) {
            $httpClient.Dispose()
        }
    }
}



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
        } else {
            throw "ProductCode not found in MSI file"
        }
        
        return [String]$value
    } catch {
        Write-Error "Error reading ProductCode from MSI file '$FilePath': $_"
        throw
    } finally {
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



function Get-MsiProperty {
    <#
    .SYNOPSIS
    Extracts a specified property from an MSI file.
    
    .DESCRIPTION
    This function uses the Windows Installer COM object to read the
    property from an MSI database file.
    
    .PARAMETER Path
    The full path to the MSI file.

    .PARAMETER PropertyName
    The name of the MSI property to retrieve.
    
    .OUTPUTS
    The specified MSI property as a string.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [String]$Path,
        [Parameter(Mandatory=$true)]
        [String]$PropertyName
    )
    
    if (-not (Test-Path $Path)) {
        throw "MSI file not found: $Path"
    }
    
    $windowsInstallerObject = $null
    $msiDatabase = $null
    $view = $null
    
    try {
        # Read property from MSI database
        $windowsInstallerObject = New-Object -ComObject WindowsInstaller.Installer
        $msiDatabase = $windowsInstallerObject.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $windowsInstallerObject, @($Path, 0))
        $query = "SELECT Value FROM Property WHERE Property = '$PropertyName'"
        $view = $msiDatabase.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $msiDatabase, ($query))
        $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)
        $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
        
        if ($record) {
            $value = $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1)
        } else {
            throw "$PropertyName not found in MSI file"
        }
        
        return [String]$value
    } catch {
        Write-Error "Error reading $PropertyName from MSI file '$Path': $_"
        throw
    } finally {
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



function Connect-AutoMSIntuneGraph {
    <#
    .SYNOPSIS
    Automatically manages Microsoft Intune Graph API connection with token refresh.
    
    .DESCRIPTION
    This function manages the connection to Microsoft Intune Graph API, automatically
    refreshing tokens when they expire or are close to expiring. It uses global
    variables for tenant configuration.
    #>
    
    if (-not $Global:TenantID -or -not $Global:ClientID -or -not $Global:ClientSecret) {
        throw "Required global variables not set: TENANT_ID, CLIENT_ID, CLIENT_SECRET"
    }
    
    # Check if the current token is invalid
    if (-not $Global:Token.ExpiresOn) {
        Write-Log "Getting an Intune Graph Client API token..."
        try {
            $Global:Token = Connect-MSIntuneGraph -TenantID $Global:TenantID -ClientID $Global:ClientID -ClientSecret $Global:ClientSecret
        } catch {
            Write-Error "Failed to get initial token: $_"
            throw
        }
    } elseif ($Global:Token.ExpiresOn.ToLocalTime() -lt (Get-Date)) {
        # If not, get a new token
        Write-Log "Token is expired. Refreshing token..."
        try {
            $Global:Token = Connect-MSIntuneGraph -TenantID $Global:TenantID -ClientID $Global:ClientID -ClientSecret $Global:ClientSecret
            Write-Log "Token refreshed. New Token Expires at: $($Global:Token.ExpiresOn.ToLocalTime())"
        } catch {
            Write-Error "Failed to refresh expired token: $_"
            throw
        }
    } elseif ($Global:Token.ExpiresOn.AddMinutes(-30).ToLocalTime() -lt (Get-Date)) {
        # For whatever reason, this API stops working 10 minutes before a token refresh
        # Set at 30 minutes in case we are uploading large files. 
        Write-Log "Token expires soon - Refreshing token..."
        try {
            # Required to force a refresh
            Clear-MsalTokenCache
            $Global:Token = Connect-MSIntuneGraph -TenantID $Global:TenantID -ClientID $Global:ClientID -ClientSecret $Global:ClientSecret
            Write-Log "Token refreshed. New Token Expires at: $($Global:Token.ExpiresOn.ToLocalTime())"
        } catch {
            Write-Error "Failed to refresh token: $_"
            throw
        }
    } else {
        Write-Log "Token is still valid. Skipping token refresh."
    } 
}



function Move-AssignmentsAndDependencies {
    <#
    .SYNOPSIS
    Moves assignments and dependencies from one Intune application to another.
    
    .DESCRIPTION
    This function transfers all group assignments and application dependencies
    from a source application to a target application, with options to offset
    availability and deadline dates.
    
    .PARAMETER From
    The source application to move assignments and dependencies from. Can be a PSObject, GUID or Display Name.
    
    .PARAMETER To
    The target application to move assignments and dependencies to. Can be a PSObject, GUID or Display Name.
    
    .PARAMETER DeadlineDateOffset
    Number of days to offset the deadline date by (default: 0).
    
    .PARAMETER AvailableDateOffset
    Number of days to offset the available date by (default: 0).
    #>
    param(
        [Parameter(Mandatory, Position=0)]
        $From,
        [Parameter(Mandatory, Position=1)]
        $To,
        [Parameter(Position=2)]
        [Int] $DeadlineDateOffset = 0,
        [Parameter(Position=3)]
        [Int] $AvailableDateOffset = 0
    )
    # If IDs or Display Names were provided instead of application objects, get the app objects
    if ($From -is [String]) {
        if ($From -match "^[0-9a-fA-F\-]{36}$") {
            # Looks like a GUID
            $From = Get-IntuneWin32App -Id $From
        }
        else {
            # Only return exact matches
            $Apps = Get-IntuneWin32App -DisplayName $From | Where-Object DisplayName -eq $From
            if ($Apps.Count -eq 1) {
                $From = $Apps[0]
            } elseif ($Apps.Count -gt 1) {
                throw "Multiple applications found with the display name '$From'. Please specify by ID instead."
            } else {
                throw "No application found with the display name '$From'."
            }
        }
    }
    if ($To -is [String]) {
        if ($To -match "[0-9a-fA-F\-]{36}$") {
            # Looks like a GUID
            $To = Get-IntuneWin32App -Id $To
        }
        else {
            # Only return exact matches
            $Apps = Get-IntuneWin32App -DisplayName $To | Where-Object DisplayName -eq $To
            if ($Apps.Count -eq 1) {
                $To = $Apps[0]
            } elseif ($Apps.Count -gt 1) {
                throw "Multiple applications found with the display name '$To'. Please specify by ID instead."
            } else {
                throw "No application found with the display name '$To'."
            }
        }
    }
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
            Write-Log "Processing assignment for group $($Assignment.GroupID) with intent $($Assignment.Intent)"
            while (!$successfullyAdded -and ($try++ -lt $maxRetries)) {
                if ($Assignment.InstallTimeSettings) {
                    Write-Log "Assignment has install time settings."
                    $useLocalTime = [bool]$Assignment.InstallTimeSettings.useLocalTime
                    Write-Log "UseLocalTime: $useLocalTime"
                    $startDateTime = $Assignment.InstallTimeSettings.startDateTime
                    $deadlineDateTime = $Assignment.InstallTimeSettings.deadlineDateTime
                    if ($null -ne $startDateTime) {
                        $startDateTime = Get-Date -Date "$AvailableDate $($startDateTime.ToString("HH:mm"))"
                        Write-Log "StartDateTime: $startDateTime"
                    }
                    if ($null -ne $deadlineDateTime) {
                        $deadlineDateTime = Get-Date -Date "$DeadlineDate $($deadlineDateTime.ToString("HH:mm"))"
                        Write-Log "DeadlineDateTime: $deadlineDateTime"
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
                else {
                    # If there is a filter
                    if ($startDateTime -and $deadlineDateTime) {
                        Write-Log "Adding assignment to $($To.id) for group $($Assignment.GroupID) with deadline and available time settings."
                        try {
                            Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $Assignment.GroupID -Intent $Assignment.Intent -Notification $Assignment.Notifications -AvailableTime $startDateTime -DeadlineTime $deadlineDateTime -UseLocalTime $useLocalTime -FilterMode $Assignment.FilterType -FilterID $Assignment.FilterID | Out-Null
                            $successfullyAdded = $true
                        } catch {
                                Write-Log "Failed to add assignment to $($To.id) for group $($Assignment.GroupID): $_"
                        }
                    } elseif ($startDateTime) {
                        Write-Log "Adding assignment to $($To.id) for group $($Assignment.GroupID) with available time settings."
                        try {
                            Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $Assignment.GroupID -Intent $Assignment.Intent -Notification $Assignment.Notifications -AvailableTime $startDateTime -UseLocalTime $useLocalTime -FilterMode $Assignment.FilterType -FilterID $Assignment.FilterID | Out-Null
                            $successfullyAdded = $true
                        } catch {
                                Write-Log "Failed to add assignment to $($To.id) for group $($Assignment.GroupID): $_"
                        }
                    } elseif ($deadlineDateTime) {
                        Write-Log "Adding assignment to $($To.id) for group $($Assignment.GroupID) with deadline time settings."
                        try {
                            Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $Assignment.GroupID -Intent $Assignment.Intent -Notification $Assignment.Notifications -DeadlineTime $deadlineDateTime -UseLocalTime $useLocalTime -FilterMode $Assignment.FilterType -FilterID $Assignment.FilterID | Out-Null
                            $successfullyAdded = $true
                        } catch {
                                Write-Log "Failed to add assignment to $($To.id) for group $($Assignment.GroupID): $_"
                        }
                    } else {
                        Write-Log "Adding assignment to $($To.id) for group $($Assignment.GroupID) without time settings."
                        try {
                                Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $Assignment.GroupID -Intent $Assignment.Intent -Notification $Assignment.Notifications -FilterMode $Assignment.FilterType -FilterID $Assignment.FilterID | Out-Null
                                $successfullyAdded = $true
                        } catch {
                                Write-Log "Failed to add assignment to $($To.id) for group $($Assignment.GroupID): $_"
                        }
                    }
                }
                if (!$successfullyAdded) {
                    Write-Log "Retrying to add assignment to $($To.id) for group $($Assignment.GroupID). Attempt $($try) of $($maxRetries)."
                    Start-Sleep -Seconds 2
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
                    } catch {
                        Write-Log "Failed to remove assignment $($Assignment.GroupID) from group $($From.id):"
                        if ($i -eq 3) {
                            Write-Log "Failed to remove assignment after 3 attempts. Skipping removal."
                            break
                        } else {
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

                Write-Log "Multiple dependencies found for $($dependency.sourceId)"
                foreach ($dependentApp in $dependentApps) {
                    $appDependenciesToMigrate = Get-IntuneWin32AppDependency -ID $dependentApp.id
                    # Clear all existing dependencies
                    Remove-IntuneWin32AppDependency -ID $dependentApp.targetId
                    # Add all dependencies back that are not the one we are moving
                    foreach ($appDependency in $appDependenciesToMigrate) {
                        $successfullyAdded = $false
                        $try = 0
                        while (!$successfullyAdded -and ($try++ -lt $maxRetries)) {
                            if ($appDependency.targetId -ne $dependency.targetId) {
                                $toAdd = New-IntuneWin32AppDependency -ID $appDependency.id -DependencyType $appDependency.dependencyType
                                Add-IntuneWin32AppDependency -ID $dependentApp.targetId -Dependency $toAdd | Out-Null
                            } else {
                                # Add the new dependency instead
                                $NewDependency = New-IntuneWin32AppDependency -ID $To.id -DependencyType $dependency.dependencyType
                                Add-IntuneWin32AppDependency -ID $dependency.targetId -Dependency $NewDependency | Out-Null
                            }
                            # Check that it worked
                            $AssignedDependencies = Get-IntuneWin32AppDependency -ID $To.id
                            if ($AssignedDependencies | Where-Object targetId -eq $dependency.targetId) {
                                $successfullyAdded = $true
                                Write-Log "Dependency $($dependency.dependencyId) added successfully to $($To.id)"
                            } else {
                                Write-Log "Retrying to add dependency $($dependency.dependencyId) to $($To.id). Attempt $($try) of $($maxRetries)."
                                Start-Sleep -Seconds 2
                            }
                        }
                    }
                }
            } elseif ($dependentApps) {
                while (!$successfullyAdded -and ($try++ -lt $maxRetries)) {
                    $NewDependency = New-IntuneWin32AppDependency -ID $dependency.targetId -DependencyType $dependency.dependencyType
                    Add-IntuneWin32AppDependency -ID $To.id -Dependency $NewDependency | Out-Null
                    # Check that it worked
                    $AssignedDependencies = Get-IntuneWin32AppDependency -ID $To.id
                    if ($AssignedDependencies | Where-Object targetId -eq $dependency.targetId) {
                        $successfullyAdded = $true
                        Write-Log "Dependency $($dependency.targetId) added successfully to $($To.id)"
                    }
                }
            } else {
                Write-Log "No dependencies found for $($dependency.sourceId). Skipping dependency move."
            }
            
        }
    }
}

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

    Write-Log "Retrieving all versions of application with display name: $DisplayName"
    
    # Attempt to connect to Intune up to 3 times
    for ($i = 0; $i -lt 3; $i++) {
        try {
            Write-Log "Attempt $($i + 1) to retrieve applications with the name $DisplayName"
            $AllSimilarApps = Get-IntuneWin32App -DisplayName "$DisplayName" -ErrorAction SilentlyContinue
            break
        }
        catch {
            Write-Log "Error retrieving applications with the name $DisplayName on attempt $($i + 1): $_"
            if ($i -eq 2) {
                Write-Log "Intune API failed to retrieve applications after 3 attempts. Exiting."
                exit 1001
            } else {
                Write-Log "Retrying to retrieve applications with the name $DisplayName. Attempt $($i + 2) of 3."
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
    # Sort by version descending, then by createdDateTime descending
    return $sortable | Sort-Object @{Expression = {[VersionPro]$_.displayVersion}; Descending = $true}, @{Expression = "createdDateTime"; Descending = $true}
}



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
        } elseif ($v1 -gt $v2) {
            return 1
        }
    }
    return 0
}



#################################################
# EMAIL NOTIFICATION FUNCTIONS
#################################################


function Initialize-ApplicationTracker {
    <#
    .SYNOPSIS
    Initializes the application tracking arrays for success/failure reporting.
    
    .DESCRIPTION
    Creates script-scoped arrays to track successful and failed application updates
    for later use in email notifications.
    #>
    $Script:SuccessfulApplications = [System.Collections.Generic.List[PSObject]]::new()
    $Script:FailedApplications = [System.Collections.Generic.List[PSObject]]::new()
}



function Add-SuccessfulApplication {
    <#
    .SYNOPSIS
    Adds an application to the successful applications tracking list.
    
    .DESCRIPTION
    Records details of a successfully processed application for inclusion
    in the summary email notification.
    
    .PARAMETER ApplicationId
    The ID of the application that was successfully processed.
    
    .PARAMETER DisplayName
    The display name of the application.
    
    .PARAMETER Version
    The version of the application that was processed.
    
    .PARAMETER Action
    The action that was performed (e.g., "Updated", "Added", "Repaired").
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$ApplicationId,
        
        [Parameter(Mandatory=$true)]
        [string]$DisplayName,
        
        [Parameter(Mandatory=$true)]
        [string]$Version,
        
        [string]$Action = "Updated"
    )
    
    if (-not $Script:SuccessfulApplications) {
        $Script:SuccessfulApplications = [System.Collections.Generic.List[PSObject]]::new()
    }
    
    $appInfo = [PSCustomObject]@{
        ApplicationId = $ApplicationId
        DisplayName = $DisplayName
        Version = $Version
        Action = $Action
        Timestamp = Get-Date
    }
    
    $Script:SuccessfulApplications.Add($appInfo)
    Write-Log "Tracked successful application: $DisplayName $Version"
}



function Add-FailedApplication {
    <#
    .SYNOPSIS
    Adds an application to the failed applications tracking list.
    
    .DESCRIPTION
    Records details of a failed application processing attempt for inclusion
    in the summary email notification.
    
    .PARAMETER ApplicationId
    The ID of the application that failed to process.
    
    .PARAMETER DisplayName
    The display name of the application.
    
    .PARAMETER Version
    The version of the application that failed to process.
    
    .PARAMETER ErrorMessage
    The error message describing why the application failed.
    
    .PARAMETER FailureStage
    The stage at which the failure occurred (e.g., "Download", "Upload", "Configuration").
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$ApplicationId,
        
        [string]$DisplayName = "Unknown",
        
        [string]$Version = "Unknown",
        
        [Parameter(Mandatory=$true)]
        [string]$ErrorMessage,
        
        [string]$FailureStage = "Processing"
    )
    
    if (-not $Script:FailedApplications) {
        $Script:FailedApplications = [System.Collections.Generic.List[PSObject]]::new()
    }
    
    $appInfo = [PSCustomObject]@{
        ApplicationId = $ApplicationId
        DisplayName = $DisplayName
        Version = $Version
        ErrorMessage = $ErrorMessage
        FailureStage = $FailureStage
        Timestamp = Get-Date
    }
    
    $Script:FailedApplications.Add($appInfo)
    Write-Log "Tracked failed application: $ApplicationId - $ErrorMessage"
}



function Test-OutlookAvailability {
    <#
    .SYNOPSIS
    Tests if Microsoft Outlook is available via COM object.
    
    .DESCRIPTION
    Checks if Outlook is installed and accessible through COM automation
    for sending email notifications.
    
    .OUTPUTS
    Boolean indicating whether Outlook is available.
    #>
    try {
        $outlook = New-Object -ComObject "Outlook.Application"
        $outlook = $null
        [System.GC]::Collect()
        return $true
    } catch {
        Write-Log "Outlook COM object not available: $_"
        return $false
    }
}



function Send-YardstickEmailReport {
    <#
    .SYNOPSIS
    Sends an email report of application processing results using Outlook COM.
    
    .DESCRIPTION
    Creates and sends an email summary of successful and failed application
    updates using Microsoft Outlook's COM interface.
    
    .PARAMETER Preferences
    Hashtable containing email configuration preferences.
    
    .PARAMETER RunParameters
    String describing the parameters used for this Yardstick run.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Preferences,
        
        [string]$RunParameters = ""
    )
    
    # Check if email notifications are enabled
    if (-not $Preferences.emailNotificationEnabled) {
        Write-Log "Email notifications are disabled in preferences"
        return
    }
    
    # Validate required email settings
    $requiredSettings = @('emailRecipient', 'emailSubject', 'emailSenderName')
    foreach ($setting in $requiredSettings) {
        if (-not $Preferences[$setting]) {
            Write-Log "WARNING: Email setting '$setting' not configured. Skipping email notification."
            return
        }
    }
    
    # Check if Outlook is available
    if (-not (Test-OutlookAvailability)) {
        Write-Log "WARNING: Outlook is not available. Cannot send email notification."
        return
    }
    
    # Initialize tracking arrays if they don't exist
    if (-not $Script:SuccessfulApplications) {
        $Script:SuccessfulApplications = [System.Collections.Generic.List[PSObject]]::new()
    }
    if (-not $Script:FailedApplications) {
        $Script:FailedApplications = [System.Collections.Generic.List[PSObject]]::new()
    }
    
    try {
        Write-Log "Creating email report for Yardstick run"
        
        # Create Outlook application and mail item
        $outlook = New-Object -ComObject "Outlook.Application"
        $mail = $outlook.CreateItem(0)  # 0 = olMailItem
        
        # Set email properties
        $mail.To = ($Preferences.emailRecipient -join "; ")

        $mail.Subject = $Preferences.emailSubject
        if ($Preferences.emailSendFromAddress) {
            $mail.SentOnBehalfOfName = $Preferences.emailSendFromAddress
        }
        
        # Build email body
        $emailBody = @"
<html>
<head>
    <style>
        body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
        .header { background-color: #0078d4; color: white; padding: 15px; border-radius: 5px 5px 0 0; }
        .content { background-color: #f8f9fa; padding: 20px; border: 1px solid #dee2e6; }
        .summary { background-color: #e7f3ff; padding: 15px; margin: 10px 0; border-radius: 5px; border-left: 4px solid #0078d4; }
        .success { background-color: #d4edda; border-left: 4px solid #28a745; }
        .failure { background-color: #f8d7da; border-left: 4px solid #dc3545; }
        .app-list { margin: 10px 0; }
        .app-item { margin: 8px 0; padding: 8px; background-color: white; border-radius: 3px; }
        .timestamp { color: #6c757d; font-size: 0.9em; }
        .error-message { color: #dc3545; font-family: monospace; margin-top: 5px; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        th, td { padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #f2f2f2; }
    </style>
</head>
<body>
    <div class="header">
        <h2>Yardstick Application Update Report</h2>
        <p>Generated by: $($Preferences.emailSenderName)</p>
        <p>Run Time: $(Get-Date -Format "MMMM dd, yyyy 'at' HH:mm:ss")</p>
$(if ($RunParameters) { "        <p>Parameters: $RunParameters</p>" })
    </div>
    
    <div class="content">
        <div class="summary">
            <h3>Summary</h3>
            <p><strong>Successful Applications:</strong> $($Script:SuccessfulApplications.Count)</p>
            <p><strong>Failed Applications:</strong> $($Script:FailedApplications.Count)</p>
            <p><strong>Total Processed:</strong> $($Script:SuccessfulApplications.Count + $Script:FailedApplications.Count)</p>
        </div>
"@

        # Add successful applications section
        if ($Script:SuccessfulApplications.Count -gt 0) {
            $emailBody += @"
        
        <div class="summary success">
            <h3>Successful Applications</h3>
            <table>
                <tr>
                    <th>Application</th>
                    <th>Version</th>
                    <th>Action</th>
                    <th>Time</th>
                </tr>
"@
            foreach ($app in $Script:SuccessfulApplications) {
                $emailBody += @"
                <tr>
                    <td><strong>$($app.DisplayName)</strong><br><small>ID: $($app.ApplicationId)</small></td>
                    <td>$($app.Version)</td>
                    <td>$($app.Action)</td>
                    <td class="timestamp">$($app.Timestamp.ToString("MM/dd/yyyy HH:mm:ss"))</td>
                </tr>
"@
            }
            $emailBody += @"
            </table>
        </div>
"@
        }

        # Add failed applications section
        if ($Script:FailedApplications.Count -gt 0) {
            $emailBody += @"
        
        <div class="summary failure">
            <h3>Failed Applications</h3>
            <table>
                <tr>
                    <th>Application</th>
                    <th>Version</th>
                    <th>Failure Stage</th>
                    <th>Error</th>
                    <th>Time</th>
                </tr>
"@
            foreach ($app in $Script:FailedApplications) {
                $emailBody += @"
                <tr>
                    <td><strong>$($app.DisplayName)</strong><br><small>ID: $($app.ApplicationId)</small></td>
                    <td>$($app.Version)</td>
                    <td>$($app.FailureStage)</td>
                    <td class="error-message">$($app.ErrorMessage)</td>
                    <td class="timestamp">$($app.Timestamp.ToString("MM/dd/yyyy HH:mm:ss"))</td>
                </tr>
"@
            }
            $emailBody += @"
            </table>
        </div>
"@
        }

        # Add footer
        $emailBody += @"
        
        <hr style="margin: 20px 0;">
        <p class="timestamp">
            <small>
                This report was automatically generated by Yardstick.<br>
                Log file location: $Global:LogLocation\$Global:LogFile
            </small>
        </p>
    </div>
</body>
</html>
"@

        # Set email body and send
        $mail.HTMLBody = $emailBody
        $mail.Send()
        
        Write-Log "Email report sent successfully to the following email addresses:"
        Write-Log ($Preferences.emailRecipient -join ", ")

        # Clean up COM objects
        $mail = $null
        $outlook = $null
        [System.GC]::Collect()
    } catch {
        Write-Log "ERROR: Failed to send email report: $_"
        # Clean up COM objects even on error
        try {
            $mail = $null
            $outlook = $null
            [System.GC]::Collect()
        } catch { }
    }
}

function Get-Secrets {
    <#
    .SYNOPSIS
    Retrieves secrets from the secrets directory
    
    .DESCRIPTION
    Connects to the secrets directory and retrieves secrets for use in the script.
    
    .PARAMETER VaultName
    The name of the vault to connect to.
    
    .OUTPUTS
    Hashtable of retrieved secrets.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$VaultName,
        [string]$SecretsDir = "$PSScriptRoot\..\Secrets"
    )
    
    # Placeholder for secret retrieval logic
    # Implement actual secret retrieval from your secure vault here
    Write-Log "Retrieving secrets from vault here: $($SecretsDir)\$VaultName.yaml"
    $Secrets = Get-Content -Path "$($SecretsDir)\$VaultName.yaml" | ConvertFrom-Yaml
    
    Write-Log "Retrieved secrets from vault: $VaultName"
    
    return $Secrets
}


# Export all functions for module availability
Export-ModuleMember -Function *
