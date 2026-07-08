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



function Test-Prerequisites {
    <#
    .SYNOPSIS
    Validates that all required prerequisites are available.

    .DESCRIPTION
    Checks for required PowerShell modules, external tools, and .NET types.
    Returns a result object with errors (fatal) and warnings (non-fatal).

    .PARAMETER ToolsPath
    Path to the Tools directory (for checking curl.exe, etc.)

    .OUTPUTS
    PSCustomObject with:
      - IsValid ([bool]) - $true if all required checks pass
      - Errors ([string[]])  - Fatal prerequisite failures
      - Warnings ([string[]]) - Non-fatal prerequisite warnings
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ToolsPath
    )

    $errors = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()

    # Required PowerShell modules (fatal if missing)
    $requiredModules = @('powershell-yaml', 'IntuneWin32App')
    foreach ($mod in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $mod)) {
            $errors.Add("Required PowerShell module '$mod' is not installed. Install with: Install-Module $mod")
        }
    }

    # Optional PowerShell modules (warn only -- needed for Adobe/SSO recipes)
    $optionalModules = @('Selenium', 'TUN.CredentialManager')
    foreach ($mod in $optionalModules) {
        if (-not (Get-Module -ListAvailable -Name $mod)) {
            $warnings.Add("Optional PowerShell module '$mod' is not installed. Adobe/SSO recipes will not work without it.")
        }
    }

    # External tools
    $curlPath = Join-Path $ToolsPath "curl.exe"
    if (-not (Test-Path $curlPath)) {
        $warnings.Add("curl.exe not found at '$curlPath'. Some download scripts may fail.")
    }

    # .NET types
    try {
        Add-Type -AssemblyName System.Net.Http -ErrorAction Stop
        [void][System.Net.Http.HttpClient]
    } catch {
        $errors.Add(".NET type System.Net.Http.HttpClient is not available. URL redirect resolution will fail.")
    }

    # COM objects
    try {
        $testInstaller = New-Object -ComObject WindowsInstaller.Installer
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($testInstaller)
    } catch {
        $warnings.Add("COM object WindowsInstaller.Installer is not available. MSI product code extraction will fail.")
    }

    return [PSCustomObject]@{
        IsValid  = ($errors.Count -eq 0)
        Errors   = [string[]]$errors
        Warnings = [string[]]$warnings
    }
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

    .PARAMETER Force
    Switch to force token refresh regardless of current token state.
    #>
    [CmdletBinding()]
    param(
        [switch]$Force
    )
    
    if (-not $Global:TenantID -or -not $Global:ClientID -or -not $Global:ClientSecret) {
        throw "Required global variables not set: TENANT_ID, CLIENT_ID, CLIENT_SECRET"
    }
    
    # If Force is specified, bypass all checks and refresh immediately
    if ($Force) {
        Write-Log "Force flag specified - Refreshing token..."
        try {
            Clear-MsalTokenCache
            $Global:Token = Connect-MSIntuneGraph -TenantID $Global:TenantID -ClientID $Global:ClientID -ClientSecret $Global:ClientSecret
            Write-Log "Token refreshed. New Token Expires at: $($Global:Token.ExpiresOn.ToLocalTime())"
        } catch {
            Write-Error "Failed to refresh token: $_"
            throw
        }
        return
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



function Invoke-WithRetry {
    <#
    .SYNOPSIS
    Invokes a script block with configurable retry logic.

    .DESCRIPTION
    Executes a script block with retry support, optional verification,
    optional rollback on failure, and optional timeout deadline.

    .PARAMETER ScriptBlock
    The script block to execute.

    .PARAMETER VerifyBlock
    Optional script block to verify success after ScriptBlock completes
    without throwing. Should return $true for success, $false for retry.

    .PARAMETER OnFailure
    Optional script block to execute when all retries are exhausted or
    timeout is reached.

    .PARAMETER MaxRetries
    Maximum number of attempts (default: 3).

    .PARAMETER DelaySeconds
    Seconds to wait between retries (default: 2).

    .PARAMETER TimeoutSeconds
    Total timeout in seconds. 0 means no timeout (default: 0).

    .PARAMETER Label
    Descriptive label for log messages (default: "operation").

    .OUTPUTS
    Returns the output of ScriptBlock on success, or $null on failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ScriptBlock]$ScriptBlock,

        [ScriptBlock]$VerifyBlock,

        [ScriptBlock]$OnFailure,

        [int]$MaxRetries = 3,

        [int]$DelaySeconds = 2,

        [int]$TimeoutSeconds = 0,

        [string]$Label = "operation"
    )

    $deadline = if ($TimeoutSeconds -gt 0) { (Get-Date).AddSeconds($TimeoutSeconds) } else { [datetime]::MaxValue }
    $lastError = $null

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        if ((Get-Date) -gt $deadline) {
            $lastError = "Timed out after $TimeoutSeconds seconds"
            Write-Log "[$Label] $lastError"
            break
        }

        try {
            $result = & $ScriptBlock

            if ($VerifyBlock) {
                $verified = & $VerifyBlock
                if (-not $verified) {
                    $lastError = "Verification failed"
                    Write-Log "[$Label] Verification failed on attempt $attempt of $MaxRetries."
                    if ($attempt -lt $MaxRetries -and (Get-Date) -lt $deadline) {
                        Start-Sleep -Seconds $DelaySeconds
                    }
                    continue
                }
            }

            return $result
        }
        catch {
            $lastError = $_.Exception.Message
            Write-Log "[$Label] Failed on attempt $attempt of ${MaxRetries}: $lastError"

            if ($attempt -lt $MaxRetries -and (Get-Date) -lt $deadline) {
                Start-Sleep -Seconds $DelaySeconds
            }
        }
    }

    Write-Log "[$Label] All $MaxRetries attempts failed. Last error: $lastError"
    if ($OnFailure) {
        try {
            & $OnFailure
        }
        catch {
            Write-Log "[$Label] OnFailure callback also failed: $_"
        }
    }

    return $null
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

    .PARAMETER AllowDependentLinkUpdates
    Boolean to allow updating dependent application links (default: $true).

    .PARAMETER DependentLinkOptions
    Hashtable of options for dependent link updates (Enabled, RetryCount, RetryDelaySeconds, TimeoutSeconds, Blacklist).

    .PARAMETER DependentUpdateStatus
    IDictionary to record the status of dependent application updates.

    .PARAMETER ProtectedSourceIds
    HashSet to record source application IDs that should be protected from deletion.
    #>
    param(
        [Parameter(Mandatory, Position=0)]
        $From,
        [Parameter(Mandatory, Position=1)]
        $To,
        [Parameter(Position=2)]
        [Int] $DeadlineDateOffset = 0,
        [Parameter(Position=3)]
        [Int] $AvailableDateOffset = 0,
        [bool] $AllowDependentLinkUpdates = $true,
        [hashtable] $DependentLinkOptions,
        [System.Collections.IDictionary] $DependentUpdateStatus,
        [System.Collections.Generic.HashSet[string]] $ProtectedSourceIds
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
    $childDependencies = @()
    $parentDependencies = @()
    if ($FromDependencies) {
        foreach ($dependency in $FromDependencies) {
            if (($dependency.PSObject.Properties.Name -contains "targetType") -and ($dependency.targetType -eq "parent")) {
                $parentDependencies += $dependency
            } else {
                $childDependencies += $dependency
            }
        }
    }

    $resolvedDependentOptions = @{
        Enabled = $true
        RetryCount = 3
        RetryDelaySeconds = 5
        TimeoutSeconds = 60
        Blacklist = @()
    }
    if ($DependentLinkOptions) {
        foreach ($key in $DependentLinkOptions.Keys) {
            if ($null -ne $DependentLinkOptions[$key]) {
                $resolvedDependentOptions[$key] = $DependentLinkOptions[$key]
            }
        }
    }
    if (-not ($resolvedDependentOptions.Blacklist -is [System.Collections.IEnumerable])) {
        $resolvedDependentOptions.Blacklist = @($resolvedDependentOptions.Blacklist)
    }
    $normalizedBlacklist = @()
    if ($resolvedDependentOptions.Blacklist) {
        $normalizedBlacklist = @($resolvedDependentOptions.Blacklist | ForEach-Object { $_.ToString().ToLowerInvariant() })
    }

    $recordDependentStatus = {
        param($name, $status)
        if ($DependentUpdateStatus -and $name) {
            $DependentUpdateStatus[$name] = $status
        }
    }
    $addProtectedSourceId = {
        param($id)
        if ($ProtectedSourceIds -and $id) {
            [void]$ProtectedSourceIds.Add($id)
        }
    }
    $normalizeDependencyType = {
        param($type)
        if (-not $type) { return "Detect" }
        switch ($type.ToString().ToLower()) {
            "autoinstall" { return "AutoInstall" }
            default { return "Detect" }
        }
    }
    if ($FromAssignments) {
        foreach ($Assignment in $FromAssignments) {
            # Skip assignments without a GroupID
            if (!$Assignment.GroupID) { 
                Write-Log "Skipping assignment without a GroupID."
                continue 
            }
            $maxRetries = 3
            $try = 0
            $successfullyAdded = $false
            Write-Verbose $Assignment
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
    if ($childDependencies -and $childDependencies.Count -gt 0) {
        foreach ($dependency in $childDependencies) {
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

    if ($parentDependencies -and $parentDependencies.Count -gt 0) {
        $parentGroups = $parentDependencies | Group-Object -Property targetId
        if (-not $AllowDependentLinkUpdates) {
            Write-Log "Skipping dependent app link updates because AllowDependentLinkUpdates is disabled for this recipe."
            foreach ($group in $parentGroups) {
                $displayName = $group.Group[0].targetDisplayName
                if (-not $displayName) { $displayName = $group.Name }
                & $recordDependentStatus $displayName "Skipped (link updates disabled for recipe)"
            }
            & $addProtectedSourceId $From.id
        }
        elseif (-not [bool]$resolvedDependentOptions.Enabled) {
            Write-Log "Skipping dependent app link updates because dependentLinkUpdateEnabled is disabled in preferences."
            foreach ($group in $parentGroups) {
                $displayName = $group.Group[0].targetDisplayName
                if (-not $displayName) { $displayName = $group.Name }
                & $recordDependentStatus $displayName "Skipped (link updates disabled in preferences)"
            }
            & $addProtectedSourceId $From.id
        }
        else {
            $maxAttempts = [int]([Math]::Max(1, $resolvedDependentOptions.RetryCount))
            $retryDelay = [int]([Math]::Max(1, $resolvedDependentOptions.RetryDelaySeconds))
            $timeoutSeconds = [int]([Math]::Max(0, $resolvedDependentOptions.TimeoutSeconds))

            foreach ($parentGroup in $parentGroups) {
            $parentId = $parentGroup.Name
            $parentDisplayName = $parentGroup.Group[0].targetDisplayName
            if (-not $parentDisplayName) {
                try {
                    $parentApp = Get-IntuneWin32App -Id $parentId
                    $parentDisplayName = $parentApp.DisplayName
                } catch {
                    Write-Log "Unable to retrieve metadata for dependent app $($parentId): $_"
                }
            }
            if (-not $parentDisplayName) {
                $parentDisplayName = $parentId
            }

            $normalizedParentName = $parentDisplayName.ToLowerInvariant()
            if ($normalizedBlacklist -and $normalizedBlacklist -contains $normalizedParentName) {
                Write-Log "Skipping dependent app $parentDisplayName because it is included in the blacklist."
                & $recordDependentStatus $parentDisplayName "Skipped (blacklisted)"
                & $addProtectedSourceId $From.id
                continue
            }

            $updateSucceeded = $false
            $statusMessage = "Updated"
            $lastError = $null
            $deadline = if ($timeoutSeconds -gt 0) { (Get-Date).AddSeconds($timeoutSeconds) } else { [datetime]::MaxValue }

            for ($attempt = 0; ($attempt -lt $maxAttempts) -and (-not $updateSucceeded); $attempt++) {
                if ((Get-Date) -gt $deadline) {
                    $lastError = "Timed out after $timeoutSeconds seconds"
                    break
                }

                $originalDependencies = @()
                $updatedDependencies = @()
                $dependenciesCleared = $false
                try {
                    $parentDependencyList = Get-IntuneWin32AppDependency -ID $parentId
                    $childItems = $parentDependencyList | Where-Object { ($_.targetType -eq "child") -or (-not $_.targetType) }
                    if (-not $childItems) {
                        $statusMessage = "No dependencies to update"
                        $updateSucceeded = $true
                        break
                    }

                    $hasLinkToSource = $false
                    foreach ($entry in $childItems) {
                        $targetAppId = $entry.targetId
                        $normalizedTypeValue = & $normalizeDependencyType $entry.dependencyType
                        $originalDependencies += New-IntuneWin32AppDependency -ID $targetAppId -DependencyType $normalizedTypeValue
                        if ($targetAppId -eq $From.id) {
                            $hasLinkToSource = $true
                            $targetAppId = $To.id
                        }
                        $updatedDependencies += New-IntuneWin32AppDependency -ID $targetAppId -DependencyType $normalizedTypeValue
                    }

                    if (-not $hasLinkToSource) {
                        $statusMessage = "Already up-to-date"
                        $updateSucceeded = $true
                        break
                    }

                    Remove-IntuneWin32AppDependency -ID $parentId | Out-Null
                    $dependenciesCleared = $true
                    Add-IntuneWin32AppDependency -ID $parentId -Dependency $updatedDependencies | Out-Null
                    $updateSucceeded = $true
                    $newTargetName = if ($To.DisplayName) { $To.DisplayName } else { $To.id }
                    Write-Log "Updated dependent app $parentDisplayName to reference $newTargetName."
                }
                catch {
                    $lastError = $_.Exception.Message
                    Write-Log "Failed to update dependent app $parentDisplayName on attempt $($attempt + 1): $lastError"
                    if ($dependenciesCleared -and $originalDependencies.Count -gt 0) {
                        try {
                            Add-IntuneWin32AppDependency -ID $parentId -Dependency $originalDependencies | Out-Null
                        } catch {
                            Write-Log "Unable to restore original dependencies for $parentDisplayName after failure."
                        }
                    }

                    if (($attempt -lt $maxAttempts - 1) -and ((Get-Date) -lt $deadline)) {
                        Start-Sleep -Seconds $retryDelay
                    }
                    else {
                        break
                    }
                }
            }

                if ($updateSucceeded) {
                    & $recordDependentStatus $parentDisplayName $statusMessage
                }
                else {
                    $failureStatus = if ($lastError) { "Failed ($lastError)" } else { "Failed" }
                    & $recordDependentStatus $parentDisplayName $failureStatus
                    & $addProtectedSourceId $From.id
                }
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

    # Attempt to retrieve applications with retry logic
    $AllSimilarApps = Invoke-WithRetry -Label "Retrieve applications for $DisplayName" -MaxRetries 3 -DelaySeconds 5 -ScriptBlock {
        Get-IntuneWin32App -DisplayName "$DisplayName" -ErrorAction Stop
    } -OnFailure {
        Write-Log "Intune API failed to retrieve applications after 3 attempts. Exiting."
        exit 1001
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



function Merge-RecipeWithBase {
    <#
    .SYNOPSIS
    Merges a recipe with its base recipe when a 'base' field is present.

    .DESCRIPTION
    Supports single-level recipe inheritance. When a recipe contains a 'base' field,
    the referenced base recipe is loaded and the child recipe's fields are overlaid
    on top. Chained inheritance (base recipe also having a 'base' field) is not supported.

    .PARAMETER Recipe
    The child recipe hashtable (already parsed from YAML).

    .PARAMETER RecipesPath
    The root path to search for the base recipe file.

    .OUTPUTS
    A merged hashtable with base fields plus child overrides, or the original recipe if no 'base' field.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Recipe,
        [Parameter(Mandatory)][string]$RecipesPath
    )

    if (-not $Recipe.ContainsKey('base')) { return $Recipe }

    $baseId = $Recipe['base']
    $baseFiles = @(Get-ChildItem $RecipesPath -Force -Recurse |
        Where-Object Name -ne 'Disabled' |
        Get-ChildItem -File -Recurse |
        Where-Object Name -match "^$baseId\.ya{0,1}ml")
    $baseFile = if ($baseFiles.Count -gt 0) { $baseFiles[0].FullName } else { $null }

    if (-not $baseFile) {
        throw "Base recipe '$baseId' not found."
    }

    $baseRecipe = Get-Content $baseFile | ConvertFrom-Yaml

    if ($baseRecipe.ContainsKey('base')) {
        throw "Chained inheritance is not supported: base recipe '$baseId' also has a 'base' field."
    }

    # Shallow merge: child fields override base fields
    $merged = $baseRecipe.Clone()
    foreach ($key in $Recipe.Keys) {
        if ($key -ne 'base') {
            $merged[$key] = $Recipe[$key]
        }
    }
    return $merged
}



function Test-RecipeSchema {
    <#
    .SYNOPSIS
    Validates a recipe hashtable against the expected schema.

    .DESCRIPTION
    Checks for required fields, conditionally required fields based on
    detectionType, valid enumeration values, and warns on unknown fields.

    .PARAMETER Recipe
    The hashtable loaded from a recipe YAML file.

    .PARAMETER RecipeId
    The application ID (filename) for error messaging.

    .OUTPUTS
    PSCustomObject with properties:
      - IsValid ([bool])
      - Errors ([string[]])
      - Warnings ([string[]])
    #>
    param(
        [Parameter(Mandatory)]
        [hashtable]$Recipe,
        [Parameter(Mandatory)]
        [string]$RecipeId
    )

    $errors = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()

    # Build combined script content for variable-in-script fallback checks
    $scriptContent = @(
        $Recipe['preDownloadScript']
        $Recipe['downloadScript']
        $Recipe['postDownloadScript']
    ) -join "`n"

    # Always-required fields
    $requiredFields = @('id', 'displayName', 'detectionType', 'iconFile', 'description', 'publisher')
    foreach ($field in $requiredFields) {
        if (-not $Recipe.ContainsKey($field) -or [string]::IsNullOrWhiteSpace($Recipe[$field])) {
            # Check if the field is set as a variable in any of the script fields
            if (-not ($scriptContent -match ('\$' + [regex]::Escape($field) + '\s*='))) {
                $errors.Add("Missing required field '$field'")
            }
        }
    }

    # Install script: at least one of installScript / powerShellInstallScript (YAML or script variable)
    if (-not $Recipe.ContainsKey('installScript') -and -not $Recipe.ContainsKey('powerShellInstallScript')) {
        if (-not ($scriptContent -match '\$installScript\s*=') -and -not ($scriptContent -match '\$powerShellInstallScript\s*=')) {
            $errors.Add("Missing install script: provide either 'installScript' or 'powerShellInstallScript'")
        }
    }

    # Uninstall script: at least one of uninstallScript / powerShellUninstallScript (YAML or script variable)
    if (-not $Recipe.ContainsKey('uninstallScript') -and -not $Recipe.ContainsKey('powerShellUninstallScript')) {
        if (-not ($scriptContent -match '\$uninstallScript\s*=') -and -not ($scriptContent -match '\$powerShellUninstallScript\s*=')) {
            $errors.Add("Missing uninstall script: provide either 'uninstallScript' or 'powerShellUninstallScript'")
        }
    }

    # Enumeration validation
    $validEnumerations = @{
        'detectionType'           = @('msi', 'file', 'script', 'registry')
        'installExperience'       = @('system', 'user')
        'restartBehavior'         = @('allow', 'suppress', 'force', 'basedOnReturnCode')
        'fileDetectionMethod'     = @('version', 'exists', 'modified', 'created', 'size')
        'registryDetectionMethod' = @('exists', 'notExists', 'string', 'integer', 'version')
    }

    foreach ($field in $validEnumerations.Keys) {
        if ($Recipe.ContainsKey($field) -and -not [string]::IsNullOrWhiteSpace($Recipe[$field])) {
            $value = $Recipe[$field].ToString().ToLower()
            $allowed = $validEnumerations[$field]
            if ($value -notin $allowed) {
                $errors.Add("Invalid value '$($Recipe[$field])' for '$field'. Allowed values: $($allowed -join ', ')")
            }
        }
    }

    # Conditionally required fields based on detectionType
    if ($Recipe.ContainsKey('detectionType') -and -not [string]::IsNullOrWhiteSpace($Recipe['detectionType'])) {
        $detType = $Recipe['detectionType'].ToString().ToLower()
        $detectionTypeFields = @{
            'file'     = @('fileDetectionPath', 'fileDetectionMethod', 'fileDetectionName')
            'registry' = @('registryDetectionMethod', 'registryDetectionKey')
            'script'   = @('detectionScript')
            'msi'      = @()
        }

        if ($detectionTypeFields.ContainsKey($detType)) {
            foreach ($field in $detectionTypeFields[$detType]) {
                if (-not $Recipe.ContainsKey($field) -or [string]::IsNullOrWhiteSpace($Recipe[$field])) {
                    # Check if the field is set as a variable in any of the script fields
                    if (-not ($scriptContent -match ('\$' + [regex]::Escape($field) + '\s*='))) {
                        $errors.Add("Missing required field '$field' for detectionType '$detType'")
                    }
                }
            }
        }
    }

    # Unknown field warnings (case-insensitive)
    $knownFields = @(
        'url', 'urlRedirects', 'id', 'version', 'fileDetectionVersion', 'displayName',
        'displayVersion', 'fileName', 'fileDetectionPath', 'preDownloadScript',
        'downloadScript', 'postDownloadScript', 'postRunScript', 'installScript',
        'uninstallScript', 'powerShellInstallScript', 'powerShellUninstallScript',
        'scopeTags', 'owner', 'maximumInstallationTimeInMinutes', 'minOSVersion',
        'installExperience', 'restartBehavior', 'availableGroups', 'requiredGroups',
        'defaultDeploymentGroups', 'allowUserUninstall', 'is32BitApp', 'architecture',
        'deadlineDateOffset', 'availableDateOffset', 'allowDependentLinkUpdates',
        'detectionType', 'fileDetectionMethod', 'fileDetectionName',
        'fileDetectionOperator', 'fileDetectionDateTime', 'fileDetectionValue',
        'registryDetectionMethod', 'registryDetectionKey', 'registryDetectionValueName',
        'registryDetectionValue', 'registryDetectionOperator', 'detectionScript',
        'detectionScriptFileExtension', 'detectionScriptRunAs32Bit',
        'detectionScriptEnforceSignatureCheck', 'iconFile', 'description', 'publisher',
        'versionLock', 'numVersionsToKeep', 'fileType', 'softwareName',
        'dependentApplicationBlacklist', 'dependentLinkUpdateEnabled',
        'dependentLinkUpdateRetryCount', 'dependentLinkUpdateRetryDelaySeconds',
        'dependentLinkUpdateTimeoutSeconds',
        'base'
    )
    $knownFieldsLower = $knownFields | ForEach-Object { $_.ToLower() }

    foreach ($key in $Recipe.Keys) {
        if ($key.ToLower() -notin $knownFieldsLower) {
            $warnings.Add("Unknown field '$key' - this field is not used by Yardstick and may be a typo")
        }
        # Check for case mismatches (key exists in known fields but with different casing)
        elseif ($key -cnotin $knownFields -and $key.ToLower() -in $knownFieldsLower) {
            $expectedCasing = $knownFields | Where-Object { $_.ToLower() -eq $key.ToLower() } | Select-Object -First 1
            $warnings.Add("Field '$key' has incorrect casing - expected '$expectedCasing'")
        }
    }

    return [PSCustomObject]@{
        IsValid  = ($errors.Count -eq 0)
        Errors   = [string[]]$errors
        Warnings = [string[]]$warnings
    }
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




# Test-VersionExcluded
# Returns whether a version falls outside the allowed version lock pattern
function Test-VersionExcluded {
    <#
    .SYNOPSIS
    Tests whether a version falls outside the allowed version lock pattern.

    .DESCRIPTION
    Returns $true if the version does NOT match the lock pattern, meaning the
    version is excluded and the update should be skipped. Returns $false if
    the version matches (is allowed) or if no lock pattern is set.
    The pattern can use 'x' as wildcards (e.g., "1.2.x" allows any 1.2.* version).

    .PARAMETER Version
    The version to check against the lock pattern.

    .PARAMETER VersionLock
    The version lock pattern. Use 'x' for wildcards.

    .OUTPUTS
    $true if the version is excluded by the lock, $false if it is allowed.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [AllowNull()]
        [AllowEmptyString()]
        [String]$Version,
        [Parameter(Mandatory=$false)]
        [String]$VersionLock
    )

    # Guard: null/empty version — treat as excluded to safely skip the update
    if ([string]::IsNullOrWhiteSpace($Version)) {
        Write-Log "WARNING: Version is null or empty in Test-VersionExcluded check"
        return $true
    }

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
        [AllowNull()]
        [AllowEmptyString()]
        [String]$Version1,
        [Parameter(Mandatory=$true)]
        [AllowNull()]
        [AllowEmptyString()]
        [String]$Version2
    )

    # Guard against null or empty version strings
    if ([string]::IsNullOrWhiteSpace($Version1)) {
        throw "Version1 is null or empty - cannot compare versions"
    }
    if ([string]::IsNullOrWhiteSpace($Version2)) {
        throw "Version2 is null or empty - cannot compare versions"
    }

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



function Test-ExtractedVersion {
    <#
    .SYNOPSIS
    Validates a version string extracted by a recipe's preDownloadScript.

    .DESCRIPTION
    Performs multiple validation checks on a version string to catch common
    scraping failures: null/empty values, HTML contamination, format violations,
    and suspiciously large jumps from the current Intune version.

    .PARAMETER Version
    The version string to validate.

    .PARAMETER ApplicationId
    The application ID for error messaging.

    .PARAMETER ExistingVersion
    Optional. The current version in Intune for major-version-jump detection.

    .OUTPUTS
    PSCustomObject with properties:
      - IsValid ([bool])
      - Errors ([string[]])
      - Warnings ([string[]])
    #>
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Version,

        [Parameter(Mandatory)]
        [string]$ApplicationId,

        [string]$ExistingVersion
    )

    $errors = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()

    # Check 1: Null or empty
    if ([string]::IsNullOrWhiteSpace($Version)) {
        $errors.Add("Version is null or empty after preDownloadScript execution")
        return [PSCustomObject]@{
            IsValid  = $false
            Errors   = [string[]]$errors
            Warnings = [string[]]$warnings
        }
    }

    # Check 2: HTML/XML contamination
    if ($Version -match '[<>]') {
        $errors.Add("Version contains HTML/XML characters: '$Version'")
    }

    # Check 3: Excessive length
    if ($Version.Length -gt 40) {
        $errors.Add("Version string is suspiciously long ($($Version.Length) chars): '$($Version.Substring(0, 40))...'")
    }

    # Check 4: Must contain at least one digit
    if ($Version -notmatch '\d') {
        $errors.Add("Version contains no digits: '$Version'")
    }

    # Check 5: Invalid characters for a version string
    if ($Version -match '[{}\[\]()=;:\"\\/@!#\$%\^&\*\|~`]') {
        $errors.Add("Version contains invalid characters: '$Version'")
    }

    # Check 6: Whitespace contamination
    if ($Version -ne $Version.Trim() -or $Version -match '[\r\n]') {
        $warnings.Add("Version contains leading/trailing whitespace or newlines: '$Version'")
    }

    # Check 7: Major version jump detection
    if ($ExistingVersion -and $errors.Count -eq 0) {
        try {
            $newClean = $Version -replace '[^0-9.]', ''
            $existClean = $ExistingVersion -replace '[^0-9.]', ''
            $newMajor = [int]($newClean.Split('.')[0])
            $existMajor = [int]($existClean.Split('.')[0])

            if ($existMajor -gt 0 -and $newMajor -gt 0 -and $newMajor -lt ($existMajor / 2)) {
                $warnings.Add("Major version dropped significantly: existing=$ExistingVersion, extracted=$Version")
            }
        } catch {
            # If we can't parse for comparison, don't block on it
        }
    }

    return [PSCustomObject]@{
        IsValid  = ($errors.Count -eq 0)
        Errors   = [string[]]$errors
        Warnings = [string[]]$warnings
    }
}



#################################################
# EMAIL NOTIFICATION FUNCTIONS
#################################################


function Format-CodeBlockHtml {
    <#
    .SYNOPSIS
    Renders a string as a scrollable, theme-aware HTML code block for the email report.

    .DESCRIPTION
    Produces a table-wrapped code block that:
      - HTML-encodes the supplied text safely.
      - Defaults to a light theme using inline styles + a bgcolor attribute, so the
        block stays legible in Outlook desktop (which ignores most CSS backgrounds).
      - Pairs with a prefers-color-scheme media query in the email's <style> block
        to switch to a dark theme on clients that report dark mode.
      - Uses a <pre> for whitespace preservation while allowing long tokens to wrap,
        so URLs and JSON payloads do not stretch the report table.

    .PARAMETER Text
    The raw error/diagnostic text to render. May contain newlines and HTML-special characters.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Text
    )

    $encoded = [System.Net.WebUtility]::HtmlEncode($Text)
    # Note: we intentionally do NOT set bgcolor attributes here. The bgcolor
    # attribute is protected from Outlook's dark-mode auto-inversion, but the
    # color on text is not — that combination produces dark-on-dark in Outlook
    # dark mode. Using inline CSS for the background lets Outlook either invert
    # both or leave both, keeping the contrast intact in either mode.
    return @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" class="code-box" style="background-color:#f5f5f5; border:1px solid #cccccc; border-radius:3px; margin:4px 0; width:100%; max-width:100%; border-collapse:separate; table-layout:fixed;"><tr><td class="code-box-cell" style="background-color:#f5f5f5; padding:8px 10px;"><div class="code-box-scroll" style="max-height:200px; max-width:100%; overflow-x:auto; overflow-y:auto;"><pre class="code-box-pre" style="display:block; margin:0; padding:0; white-space:pre; font-family:Consolas,'Courier New',monospace; font-size:12px; color:#1a1a1a; background:transparent;">$encoded</pre></div></td></tr></table>
"@
}



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
    
    .PARAMETER Dependents
    A hashtable containing dependent app names and their status (e.g., "Added", "Not Updated (Auto-update disabled)", "Failed").
    
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

        # Store a list of dependent apps for this application and whether or not they had their references updated successfully
        [hashtable]$Dependents = @{},

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
        Dependents = $Dependents
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

    # Check if app with same ID already exists in failed list
    if (-not ($Script:FailedApplications | Where-Object { $_.ApplicationId -eq $ApplicationId })) {
        $Script:FailedApplications.Add($appInfo)
        Write-Log "Tracked failed application: $ApplicationId - $ErrorMessage"
    } else {
        Write-Log "Application $ApplicationId already exists in failed applications list. Skipping duplicate entry."
    }
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

        [string]$RunParameters = "",

        # When set, the email is opened in Outlook for visual inspection instead of being sent.
        [switch]$Preview,

        # Optional path. If provided, the rendered HTML body is also written to this file
        # so it can be opened in a browser (useful when validating formatting changes).
        [string]$HtmlOutputPath
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
        
        # Load branding logo. Use a cid: reference for the email itself (Outlook
        # desktop does not render base64 data URIs reliably) and a base64 data URI
        # for the HtmlOutputPath so the standalone file stays self-contained.
        $logoHtml = ""
        $logoHtmlForBrowser = ""
        $logoCid = "yardstick-logo"
        $logoPath = $null
        try {
            $candidatePath = Join-Path $PSScriptRoot "..\Branding\yardstick_logo_white_text_transparent_bg.png"
            if (Test-Path $candidatePath) {
                $logoPath = (Resolve-Path $candidatePath).Path
                $logoBytes = [System.IO.File]::ReadAllBytes($logoPath)
                $logoBase64 = [System.Convert]::ToBase64String($logoBytes)
                $logoHtml = "<img src=`"cid:$logoCid`" alt=`"Yardstick`" width=`"180`" height=`"180`" class=`"header-logo`" style=`"width:180px; height:180px; max-width:100%; display:block; border:0;`" />"
                $logoHtmlForBrowser = "<img src=`"data:image/png;base64,$logoBase64`" alt=`"Yardstick`" width=`"180`" height=`"180`" class=`"header-logo`" style=`"width:180px; height:180px; max-width:100%; display:block; border:0;`" />"
            } else {
                Write-Log "WARNING: Branding logo not found at $candidatePath"
            }
        } catch {
            Write-Log "WARNING: Failed to embed branding logo: $($_.Exception.Message)"
        }

        # Build email body
        $emailBody = @"
<html>
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        body { font-family: Segoe UI, Arial, sans-serif; margin: 0; padding: 0; background-color: #e9ecef; }
        html { background-color: #e9ecef; }
        .header { background-color: #46C4DD; color: white; padding: 15px; border-radius: 5px 5px 0 0; }
        .header .logo { width: 180px; height: auto; float: left; margin-right: 15px; display: block; }
        .header-clear { clear: both; }
        .header h2 { margin: 0 0 5px 0; }
        .header p { margin: 2px 0; }
        .content { background-color: #f8f9fa; padding: 20px; border: 1px solid #dee2e6; }
        .summary { background-color: #e7f3ff; padding: 15px; margin: 10px 0; border-radius: 5px; border-left: 4px solid #0078d4; }
        .success { background-color: #d4edda; border-left: 4px solid #28a745; }
        .failure { background-color: #f8d7da; border-left: 4px solid #dc3545; }
        .app-list { margin: 10px 0; }
        .app-item { margin: 8px 0; padding: 8px; background-color: white; border-radius: 3px; }
        .timestamp { color: #6c757d; font-size: 0.9em; }
        .table-wrapper { width: 100%; max-width: 100%; overflow-x: auto; margin: 10px 0; }
        .error-message { color: #dc3545; font-family: monospace; margin-top: 5px; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; table-layout: fixed; }
        th, td { padding: 8px; text-align: left; border-bottom: 1px solid #ddd; vertical-align: top; overflow-wrap: anywhere; word-wrap: break-word; word-break: break-word; }
        th { background-color: #f2f2f2; }

        /* Reset the global td styling for code-box inner cell so it doesn't get
           an extra border or padding from the report table rules above. */
        .code-box { border-collapse: separate !important; }
        .code-box .code-box-cell { border-bottom: none !important; }

        /* Theme-aware code box: light by default (legible everywhere, including
           Outlook desktop which will auto-invert both bg and text together in
           its dark mode). Modern clients honoring prefers-color-scheme, and
           Outlook.com / Outlook iOS via [data-ogsc] / [data-ogsb], get an
           explicit dark theme. */
        @media (prefers-color-scheme: dark) {
            .code-box, .code-box .code-box-cell {
                background-color: #1e1e1e !important;
                border-color: #444 !important;
            }
            .code-box-pre { color: #f8f8f2 !important; }
        }
        [data-ogsc] .code-box,
        [data-ogsc] .code-box .code-box-cell,
        [data-ogsb] .code-box,
        [data-ogsb] .code-box .code-box-cell {
            background-color: #1e1e1e !important;
            border-color: #444 !important;
        }
        [data-ogsc] .code-box-pre,
        [data-ogsb] .code-box-pre { color: #f8f8f2 !important; }

        @media only screen and (max-width: 768px) {
            .header { table-layout: fixed !important; max-width: 100% !important; }
            .code-box-scroll {
                overflow-x: hidden !important;
            }
            .code-box-pre {
                white-space: pre-wrap !important;
                overflow-wrap: anywhere !important;
                word-wrap: break-word !important;
                word-break: break-word !important;
            }
            .header-logo {
                width: 100px !important;
                height: 100px !important;
                margin: 0 auto !important;
            }
            .header-logo-cell,
            .header-text-cell {
                display: block !important;
                width: 100% !important;
                max-width: 100% !important;
                box-sizing: border-box !important;
                text-align: center !important;
                padding: 10px 15px !important;
            }
            .report-table { display: block !important; }
            .report-table thead,
            .report-table .report-header-row { display: none !important; }
            .report-table tbody,
            .report-table tr,
            .report-table td {
                display: block !important;
                width: 100% !important;
                box-sizing: border-box !important;
            }
            .report-table tr {
                margin: 0 0 12px 0 !important;
                border: 1px solid rgba(0,0,0,0.12) !important;
                border-radius: 4px !important;
                background-color: #ffffff !important;
                color: #1a1a1a !important;
                padding: 4px 0 !important;
            }
            .report-table td {
                border: none !important;
                padding: 6px 12px !important;
                text-align: left !important;
                color: #1a1a1a !important;
            }
            .report-table td[data-label]:before {
                content: attr(data-label) ": ";
                font-weight: bold;
                color: #555 !important;
                display: inline-block;
                margin-right: 4px;
            }
            .report-table tr.dependents-row {
                margin-top: -10px !important;
                background-color: #f8f9fa !important;
                color: #1a1a1a !important;
                border-top: none !important;
                border-radius: 0 0 4px 4px !important;
            }
        }

        @media only screen and (max-width: 1280px) and (prefers-color-scheme: dark) {
            .report-table tr,
            .report-table tr.dependents-row {
                background-color: #2a2a2a !important;
                color: #f0f0f0 !important;
                border-color: #444 !important;
            }
            .report-table td { color: #f0f0f0 !important; }
            .report-table td[data-label]:before { color: #b0b0b0 !important; }
        }
        [data-ogsc] .report-table tr,
        [data-ogsb] .report-table tr,
        [data-ogsc] .report-table tr.dependents-row,
        [data-ogsb] .report-table tr.dependents-row {
            background-color: #2a2a2a !important;
            color: #f0f0f0 !important;
            border-color: #444 !important;
        }
        [data-ogsc] .report-table td,
        [data-ogsb] .report-table td { color: #f0f0f0 !important; }
        [data-ogsc] .report-table td[data-label]:before,
        [data-ogsb] .report-table td[data-label]:before { color: #b0b0b0 !important; }

        /* Section containers: keep dark text on the light pastel backgrounds in
           light mode, and explicitly flip both background + text for dark mode
           so neither half-applies. */
        @media (prefers-color-scheme: dark) {
            .section-summary { background-color: #1e2a33 !important; color: #f0f0f0 !important; }
            .section-success { background-color: #1e2e1f !important; color: #f0f0f0 !important; }
            .section-failure { background-color: #2e1e20 !important; color: #f0f0f0 !important; }
            .section-summary h3, .section-summary p, .section-summary strong,
            .section-success h3, .section-success p, .section-success strong,
            .section-failure h3, .section-failure p, .section-failure strong {
                color: #f0f0f0 !important;
            }
        }
        [data-ogsc] .section-summary,
        [data-ogsb] .section-summary { background-color: #1e2a33 !important; color: #f0f0f0 !important; }
        [data-ogsc] .section-success,
        [data-ogsb] .section-success { background-color: #1e2e1f !important; color: #f0f0f0 !important; }
        [data-ogsc] .section-failure,
        [data-ogsb] .section-failure { background-color: #2e1e20 !important; color: #f0f0f0 !important; }
        [data-ogsc] .section-summary h3, [data-ogsb] .section-summary h3,
        [data-ogsc] .section-summary p,  [data-ogsb] .section-summary p,
        [data-ogsc] .section-summary strong, [data-ogsb] .section-summary strong,
        [data-ogsc] .section-success h3, [data-ogsb] .section-success h3,
        [data-ogsc] .section-failure h3, [data-ogsb] .section-failure h3 {
            color: #f0f0f0 !important;
        }
    </style>
</head>
<body bgcolor="#e9ecef" style="background-color:#e9ecef;">
    <table class="header" cellpadding="0" cellspacing="0" border="0" style="width:100%; background-color:#0078d4; color:white; border-radius:5px; border-collapse:collapse;">
        <tr>
            <td class="header-logo-cell" style="width:195px; padding:15px 15px 15px 15px; vertical-align:middle; border:0;">$logoHtml</td>
            <td class="header-text-cell" style="padding:15px 15px 15px 0; vertical-align:middle; border:0;">
                <p style="margin:0 0 6px 0; padding:0; font-size:18pt; font-weight:bold; line-height:1.1; color:white;">Application Update Report</p>
                <p style="margin:0; padding:0; line-height:1.2; color:white;">Run Time: $(Get-Date -Format "MMMM dd, yyyy 'at' HH:mm:ss tt")</p>
$(if ($RunParameters) { @"
                <p style="margin:6px 0 4px 0; padding:0; line-height:1.4; color:white;">Parameters:</p>
                <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%; max-width:100%; table-layout:fixed; border-collapse:collapse; margin:0;">
                    <tr>
                        <td style="width:100%; padding:0;">
                            <div style="max-width:100%; overflow-x:auto; overflow-y:hidden; background-color:#f5f5f5; border:1px solid #cccccc; border-radius:3px;">
                                <code style="display:inline-block; white-space:nowrap; padding:4px 8px; background:transparent; color:#1a1a1a; font-family:Consolas,'Courier New',monospace; font-size:11pt;">$([System.Net.WebUtility]::HtmlEncode($RunParameters))</code>
                            </div>
                        </td>
                    </tr>
                </table>
"@ })
            </td>
        </tr>
    </table>

    <div class="content">
        <table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#e7f3ff" style="width:100%; background-color:#e7f3ff; border-collapse:separate; border-radius:5px; margin:10px 0;">
            <tr>
                <td class="section-summary" bgcolor="#e7f3ff" style="background-color:#e7f3ff; color:#1a1a1a; padding:15px; border-left:4px solid #0078d4; border-radius:5px;">
                    <h3 style="margin:0 0 8px 0; padding:0; line-height:1.2; color:#1a1a1a;">Summary</h3>
                    <p style="margin:0; padding:0; line-height:1.3; color:#1a1a1a;"><strong>Successful Applications:</strong> $($Script:SuccessfulApplications.Count)</p>
                    <p style="margin:0; padding:0; line-height:1.3; color:#1a1a1a;"><strong>Failed Applications:</strong> $($Script:FailedApplications.Count)</p>
                    <p style="margin:0; padding:0; line-height:1.3; color:#1a1a1a;"><strong>Total Processed:</strong> $($Script:SuccessfulApplications.Count + $Script:FailedApplications.Count)</p>
                </td>
            </tr>
        </table>
"@

        if ($Script:SuccessfulApplications.Count -gt 0) {
            $emailBody += @"

        <table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#d4edda" style="width:100%; background-color:#d4edda; border-collapse:separate; border-radius:5px; margin:10px 0;">
            <tr>
                <td class="section-success" bgcolor="#d4edda" style="background-color:#d4edda; color:#1a1a1a; padding:15px; border-left:4px solid #28a745; border-radius:5px;">
                    <h3 style="margin:0 0 8px 0; padding:0; line-height:1.2; color:#1a1a1a;">Successful Applications</h3>
                    <table class="report-table">
                        <tr class="report-header-row">
                            <th>Application</th>
                            <th>Version</th>
                            <th>Action</th>
                            <th>Time</th>
                        </tr>
"@
            foreach ($app in $Script:SuccessfulApplications) {
                $emailBody += @"
                <tr>
                    <td data-label="Application"><strong>$($app.DisplayName)</strong><br><small>ID: $($app.ApplicationId)</small></td>
                    <td data-label="Version">$($app.Version)</td>
                    <td data-label="Action">$($app.Action)</td>
                    <td data-label="Time" class="timestamp">$($app.Timestamp.ToString("MM/dd/yyyy HH:mm:ss"))</td>
                </tr>
"@
                # Add dependency information if present
                if ($app.Dependents -and $app.Dependents.Count -gt 0) {
                    $emailBody += @"
                <tr class="dependents-row">
                    <td colspan="4" style="background-color: #f8f9fa; color:#1a1a1a; padding-left: 30px;">
                        <strong style="color:#1a1a1a;">Dependent Applications:</strong>
                        <ul style="margin: 5px 0;">
"@
                    foreach ($dep in $app.Dependents.GetEnumerator()) {
                        $depValue = [string]$dep.Value
                        $depName  = [string]$dep.Key

                        $statusText  = $depValue
                        $errorDetail = $null
                        if ($depValue -match '^Failed\s*\((.+)\)\s*$') {
                            $statusText  = 'Failed'
                            $errorDetail = $matches[1]
                        }

                        $statusColor = switch -Regex ($statusText) {
                            "^Added$|^Updated$" { "#28a745" }
                            "^Not Updated|^Skipped" { "#ffc107" }
                            "^Failed" { "#dc3545" }
                            default { "#6c757d" }
                        }

                        $encodedName   = [System.Net.WebUtility]::HtmlEncode($depName)
                        $encodedStatus = [System.Net.WebUtility]::HtmlEncode($statusText)

                        if ($errorDetail) {
                            $errorBox = Format-CodeBlockHtml -Text $errorDetail
                            $emailBody += @"
                            <li style="margin-bottom: 10px;"><span style="color: $statusColor; font-weight: bold;">$encodedStatus</span> - $encodedName$errorBox</li>
"@
                        } else {
                            $emailBody += @"
                            <li><span style="color: $statusColor; font-weight: bold;">$encodedStatus</span> - $encodedName</li>
"@
                        }
                    }
                    $emailBody += @"
                        </ul>
                    </td>
                </tr>
"@
                }
            }
            $emailBody += @"
                    </table>
                </td>
            </tr>
        </table>
"@
        }
        if ($Script:FailedApplications.Count -gt 0) {
            $emailBody += @"

        <table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#f8d7da" style="width:100%; background-color:#f8d7da; border-collapse:separate; border-radius:5px; margin:10px 0;">
            <tr>
                <td class="section-failure" bgcolor="#f8d7da" style="background-color:#f8d7da; color:#1a1a1a; padding:15px; border-left:4px solid #dc3545; border-radius:5px;">
                    <h3 style="margin:0 0 8px 0; padding:0; line-height:1.2; color:#1a1a1a;">Failed Applications</h3>
                    <div class="table-wrapper">
                    <table class="report-table">
                        <tr class="report-header-row">
                            <th>Application</th>
                            <th>Version</th>
                            <th>Failure Stage</th>
                            <th>Error</th>
                            <th>Time</th>
                        </tr>
"@
            foreach ($app in $Script:FailedApplications) {
                $errorCodeBox = Format-CodeBlockHtml -Text ([string]$app.ErrorMessage)
                $emailBody += @"
                <tr>
                    <td data-label="Application"><strong>$($app.DisplayName)</strong><br><small>ID: $($app.ApplicationId)</small></td>
                    <td data-label="Version">$($app.Version)</td>
                    <td data-label="Failure Stage">$($app.FailureStage)</td>
                    <td data-label="Error">$errorCodeBox</td>
                    <td data-label="Time" class="timestamp">$($app.Timestamp.ToString("MM/dd/yyyy HH:mm:ss"))</td>
                </tr>
"@
            }
            $emailBody += @"
                    </table>
                    </div>
                </td>
            </tr>
        </table>
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

        # Set email body and send (or open for preview)
        $mail.HTMLBody = $emailBody

        # Attach the logo as an inline image with a Content-ID matching the cid:
        # reference in the HTML body.
        if ($logoPath) {
            try {
                $attachment = $mail.Attachments.Add($logoPath)
                $pa = $attachment.PropertyAccessor
                # Set the Content-ID so it can be referenced via cid: in the HTML
                $pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", $logoCid)
                # Note: PR_ATTACH_FLAGS (0x37140003) is often read-only in Outlook COM and doesn't
                # support SetProperty; attempting it would cause an error. The cid: reference still
                # works for inline rendering even if the attachment is visible in the attachment list.
            } catch {
                Write-Log "WARNING: Failed to attach logo with CID: $($_.Exception.Message)"
            }
        }

        if ($HtmlOutputPath) {
            try {
                $htmlForFile = $emailBody
                if ($logoHtmlForBrowser -and $logoHtml) {
                    $htmlForFile = $htmlForFile.Replace($logoHtml, $logoHtmlForBrowser)
                }
                Set-Content -Path $HtmlOutputPath -Value $htmlForFile -Encoding UTF8
                Write-Log "Rendered email HTML written to $HtmlOutputPath"
            } catch {
                Write-Log "WARNING: Failed to write rendered HTML to $HtmlOutputPath`: $_"
            }
        }

        if ($Preview) {
            $mail.Display()
            Write-Log "Email opened in Outlook for preview (not sent)."
        } else {
            $mail.Send()
            Write-Log "Email report sent successfully to the following email addresses:"
            Write-Log ($Preferences.emailRecipient -join ", ")
        }
        
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
