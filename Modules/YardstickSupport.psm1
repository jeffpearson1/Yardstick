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
        Write-Warning "LogLocation or LogFile variables are not set. Cannot write to log."
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

    .PARAMETER IntentFilter
    When set, only assignments matching this intent are processed (e.g. 'required',
    'available'). Empty (default) processes all intents.

    .PARAMETER CopyOnly
    When set, assignments are added to $To but not removed from $From.

    .PARAMETER SkipDependencies
    When set, dependency migration (child + parent link rewriting) is skipped.
    Useful when calling this function twice for different intents against the
    same From/To pair - dependencies only need to be migrated once.

    .NOTES
    All Devices and All Users assignments that use an assignment filter are not
    migrated. Add-IntuneWin32AppAssignmentAllDevices/AllUsers accept a filter by
    name only, while Get-IntuneWin32AppAssignment reports the filter id, so the
    filter cannot be carried across - and migrating without it would widen the
    assignment to every device or user. Those assignments are logged and left on
    $From for manual handling.

    Removal matches on the target alone (a group id, or the virtual target type)
    because the IntuneWin32App module cannot delete an individual assignment. If
    $From holds more than one assignment for the same target, migrating one of
    them removes them all; this is logged when it happens.
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
        [System.Collections.Generic.HashSet[string]] $ProtectedSourceIds,
        [ValidateSet('', 'required', 'available')]
        [string] $IntentFilter = '',
        [switch] $CopyOnly,
        [switch] $SkipDependencies
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
    # KNOWN ISSUE (IntuneWin32App 1.5.0): Get-IntuneWin32AppAssignment returns
    # $null for any app that has EXACTLY ONE assignment, so the migration below
    # silently does nothing for those apps. Get-IntuneWin32AppAssignment.ps1:124
    # reads the assignments with Invoke-MSGraphOperation, which unrolls a
    # single-element response to a bare [PSCustomObject]; the guard on the next
    # line then tests `$response.Count -gt 0`, and PSCustomObject has no
    # synthetic .Count (unlike other scalars in PS 3.0+), so it evaluates $null
    # -gt 0 = $false and the cmdlet reports "No assignments found". Two or more
    # assignments come back as an Object[] and work fine. Wrapping the call in
    # @() does not help - the data is already discarded inside the cmdlet.
    # Fixing this needs a Graph-direct reader (see Set-AssignmentAutoUpdate,
    # which already bypasses the cmdlet via Invoke-YardstickGraphRequest).
    $FromAssignments = Get-IntuneWin32AppAssignment -Id $From.id
    $FromDependencies = Get-IntuneWin32AppDependency -Id $From.id
    # Kept as DateTime, not a formatted string: rebuilding a date by formatting
    # to "MM/dd/yyyy" and parsing it back with Get-Date makes the result depend
    # on the host's culture, so on a dd/MM/yyyy host 08/04 silently becomes
    # 8 April and a day past the 12th throws outright.
    $AvailableDate = (Get-Date).AddDays($AvailableDateOffset).Date
    $DeadlineDate = (Get-Date).AddDays($DeadlineDateOffset).Date
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
    # Assignment targets that carry no GroupID and need their own cmdlets.
    $allDevicesTarget = "#microsoft.graph.allDevicesAssignmentTarget"
    $allUsersTarget = "#microsoft.graph.allLicensedUsersAssignmentTarget"

    if ($FromAssignments) {
        # Removal matches on the target alone - a group id, or the virtual target
        # type - so a source app holding two assignments for the same target (an
        # include plus an exclude, or two different intents) loses both when one
        # of them is migrated. The module exposes no way to delete a single
        # assignment, so warn rather than removing more than was migrated.
        $targetCounts = @{}
        foreach ($existingAssignment in $FromAssignments) {
            $countKey = if ($existingAssignment.GroupID) { $existingAssignment.GroupID } else { $existingAssignment.Type }
            if ($countKey) {
                $targetCounts[$countKey] = 1 + [int]$targetCounts[$countKey]
            }
        }

        foreach ($Assignment in $FromAssignments) {
            $targetType = $Assignment.Type
            $isVirtualTarget = ($targetType -eq $allDevicesTarget) -or ($targetType -eq $allUsersTarget)
            $isExclusion = ($Assignment.GroupMode -eq "Exclude")
            $hasFilter = $Assignment.FilterID -and $Assignment.FilterType -and ($Assignment.FilterType -ne "none")
            $assignmentLabel = if ($targetType -eq $allDevicesTarget) {
                "All Devices"
            }
            elseif ($targetType -eq $allUsersTarget) {
                "All Users"
            }
            elseif ($Assignment.GroupID) {
                "group $($Assignment.GroupID)"
            }
            else {
                $null
            }

            if (-not $assignmentLabel) {
                Write-Log "Skipping assignment with no GroupID and unsupported target type '$targetType'."
                continue
            }
            if ($IntentFilter -and ($Assignment.Intent -ne $IntentFilter)) {
                Write-Log "Skipping assignment for $assignmentLabel - intent $($Assignment.Intent) does not match filter $IntentFilter"
                continue
            }
            # Add-IntuneWin32AppAssignmentAllDevices/AllUsers take a filter by
            # name only, and Get-IntuneWin32AppAssignment reports the filter id,
            # so the filter cannot be carried across. Dropping it would widen the
            # assignment to every device or user, so refuse to migrate instead.
            if ($isVirtualTarget -and $hasFilter) {
                Write-Log "WARNING: Skipping $assignmentLabel assignment because filter $($Assignment.FilterID) ($($Assignment.FilterType)) cannot be reapplied by ID. Migrate this assignment manually."
                continue
            }

            $maxRetries = 3
            $try = 0
            $successfullyAdded = $false
            Write-Verbose $Assignment
            Write-Log "Processing $assignmentLabel assignment with intent $($Assignment.Intent)$(if ($isExclusion) { ' (exclusion)' })"

            # Reset per assignment - these must not leak into the next iteration.
            $startDateTime = $null
            $deadlineDateTime = $null
            $useLocalTime = $false
            if ($Assignment.InstallTimeSettings) {
                $useLocalTime = [bool]$Assignment.InstallTimeSettings.useLocalTime
                $sourceStart = $Assignment.InstallTimeSettings.startDateTime
                $sourceDeadline = $Assignment.InstallTimeSettings.deadlineDateTime
                # Rebase onto the offset date, keeping the source time of day.
                # Cast defensively: Graph hands these back as DateTime, but a
                # string would make .Hour/.Minute silently unavailable.
                if ($null -ne $sourceStart) {
                    $sourceStart = [datetime]$sourceStart
                    $startDateTime = $AvailableDate.AddHours($sourceStart.Hour).AddMinutes($sourceStart.Minute)
                }
                if ($null -ne $sourceDeadline) {
                    $sourceDeadline = [datetime]$sourceDeadline
                    $deadlineDateTime = $DeadlineDate.AddHours($sourceDeadline.Hour).AddMinutes($sourceDeadline.Minute)
                }

                # Add-IntuneWin32AppAssignmentGroup rejects a deadline that is
                # already in the past unless an available time accompanies it -
                # and it rejects it with `break`, which escapes our try/catch and
                # kills the foreach, silently abandoning every assignment still
                # to be migrated. Rebasing onto today (offset 0) puts any
                # morning deadline in the past for an afternoon run, so nudge it
                # forward. Intune treats a just-passed deadline the same way:
                # install at the next check-in.
                if (($null -ne $deadlineDateTime) -and ($null -eq $startDateTime) -and ($deadlineDateTime -lt (Get-Date))) {
                    $adjustedDeadline = (Get-Date).AddMinutes(5)
                    Write-Log "WARNING: Rebased deadline $deadlineDateTime for $assignmentLabel is in the past; moving it to $adjustedDeadline so the assignment is still accepted."
                    $deadlineDateTime = $adjustedDeadline
                }

                Write-Log "Install time settings - UseLocalTime: $useLocalTime, StartDateTime: $startDateTime, DeadlineDateTime: $deadlineDateTime"
            }

            $assignmentParams = @{
                ID     = $To.id
                Intent = $Assignment.Intent
            }
            if ($isExclusion) {
                # An exclusion carries no settings of its own: the Exclude
                # parameter set accepts only ID, GroupID and Intent.
                $assignmentParams["Exclude"] = $true
                $assignmentParams["GroupID"] = $Assignment.GroupID
            }
            else {
                if (-not $isVirtualTarget) {
                    $assignmentParams["Include"] = $true
                    $assignmentParams["GroupID"] = $Assignment.GroupID
                }
                if ($Assignment.Notifications) {
                    $assignmentParams["Notification"] = $Assignment.Notifications
                }
                if ($startDateTime) {
                    $assignmentParams["AvailableTime"] = $startDateTime
                }
                if ($deadlineDateTime) {
                    $assignmentParams["DeadlineTime"] = $deadlineDateTime
                }
                if ($startDateTime -or $deadlineDateTime) {
                    $assignmentParams["UseLocalTime"] = $useLocalTime
                }
                if ($hasFilter -and (-not $isVirtualTarget)) {
                    $assignmentParams["FilterMode"] = $Assignment.FilterType
                    $assignmentParams["FilterID"] = $Assignment.FilterID
                }
            }

            while (!$successfullyAdded -and ($try++ -lt $maxRetries)) {
                try {
                    # The IntuneWin32App cmdlets bail out of their Begin block
                    # with `break` (past deadline, expired token). A bare `break`
                    # from a called function is not catchable and unwinds to the
                    # caller's nearest enclosing loop - without this single-pass
                    # foreach to absorb it, one bad assignment would silently
                    # abandon every assignment after it. Absorbed here, it just
                    # leaves $addReturned false and is retried and logged.
                    $addReturned = $false
                    foreach ($breakGuard in 1) {
                        if ($targetType -eq $allDevicesTarget) {
                            Add-IntuneWin32AppAssignmentAllDevices @assignmentParams | Out-Null
                        }
                        elseif ($targetType -eq $allUsersTarget) {
                            Add-IntuneWin32AppAssignmentAllUsers @assignmentParams | Out-Null
                        }
                        else {
                            Add-IntuneWin32AppAssignmentGroup @assignmentParams | Out-Null
                        }
                        $addReturned = $true
                    }
                    if ($addReturned) {
                        $successfullyAdded = $true
                        Write-Log "Added $assignmentLabel assignment to $($To.id)."
                    }
                    else {
                        Write-Log "Add of $assignmentLabel assignment to $($To.id) aborted before returning. Check the preceding warning for the reason."
                    }
                }
                catch {
                    Write-Log "Failed to add $assignmentLabel assignment to $($To.id): $_"
                }
                if (!$successfullyAdded -and ($try -lt $maxRetries)) {
                    Write-Log "Retrying to add $assignmentLabel assignment to $($To.id). Attempt $($try) of $($maxRetries)."
                    Start-Sleep -Seconds 2
                }
            }

            # Remove the old assignment
            if ($successfullyAdded -and -not $CopyOnly) {
                $countKey = if ($Assignment.GroupID) { $Assignment.GroupID } else { $Assignment.Type }
                if ([int]$targetCounts[$countKey] -gt 1) {
                    Write-Log "WARNING: $($From.id) has $($targetCounts[$countKey]) assignments for $assignmentLabel. Removing one removes them all."
                }

                # The Add-IntuneWin32AppAssignment* cmdlets downgrade Graph
                # failures and duplicate-target conflicts to warnings, so a
                # returned call does not prove the assignment landed. Confirm it
                # is really on $To before deleting the only remaining copy.
                # Match on the same key used for removal so virtual targets
                # (All Devices/All Users, which carry no GroupID) compare by
                # target type instead of collapsing onto a null GroupID.
                $verified = $false
                try {
                    $verified = @(Get-IntuneWin32AppAssignment -Id $To.id | Where-Object {
                        $existingKey = if ($_.GroupID) { $_.GroupID } else { $_.Type }
                        ($existingKey -eq $countKey) -and ($_.Intent -eq $Assignment.Intent)
                    }).Count -gt 0
                }
                catch {
                    Write-Log "Could not verify $assignmentLabel assignment on $($To.id): $_"
                }
                if (-not $verified) {
                    Write-Log "ERROR: $assignmentLabel assignment was not found on $($To.id) after the add; leaving the source assignment on $($From.id) intact."
                    continue
                }

                $maxRemovalAttempts = 3
                $successfullyRemoved = $false
                for ($i = 1; ($i -le $maxRemovalAttempts) -and (-not $successfullyRemoved); $i++) {
                    $removeFailure = $null
                    try {
                        # Same single-pass foreach as above, absorbing a `break`
                        # out of the cmdlet's Begin block so it cannot unwind
                        # this retry loop and the assignment loop around it.
                        $removeReturned = $false
                        foreach ($breakGuard in 1) {
                            if ($targetType -eq $allDevicesTarget) {
                                Remove-IntuneWin32AppAssignmentAllDevices -ID $From.id | Out-Null
                            }
                            elseif ($targetType -eq $allUsersTarget) {
                                Remove-IntuneWin32AppAssignmentAllUsers -ID $From.id | Out-Null
                            }
                            else {
                                Remove-IntuneWin32AppAssignmentGroup -ID $From.id -GroupID $Assignment.GroupID | Out-Null
                            }
                            $removeReturned = $true
                        }
                        if ($removeReturned) {
                            $successfullyRemoved = $true
                            Write-Log "Successfully removed $assignmentLabel assignment from $($From.id)."
                        }
                        else {
                            $removeFailure = "aborted before returning. Check the preceding warning for the reason"
                        }
                    }
                    catch {
                        $removeFailure = $_
                    }
                    if (-not $successfullyRemoved) {
                        Write-Log "Failed to remove $assignmentLabel assignment from $($From.id): $removeFailure"
                        if ($i -eq $maxRemovalAttempts) {
                            Write-Log "Failed to remove assignment after $maxRemovalAttempts attempts. Skipping removal."
                        }
                        else {
                            Write-Log "Retrying removal of $assignmentLabel assignment from $($From.id). Attempt $($i + 1) of $maxRemovalAttempts."
                            Start-Sleep -Seconds 2
                        }
                    }
                }
            }
        }
    }
    if ($SkipDependencies) {
        Write-Log "Skipping dependency migration as requested."
        return
    }
    # Child dependencies are the apps that $From depends on; $To needs to depend
    # on the same apps. Add-IntuneWin32AppDependency REPLACES an app's entire
    # dependency set (it only preserves supersedence), so the complete desired
    # list has to be submitted in one call - adding them one at a time drops
    # every dependency configured by the previous call.
    if ($childDependencies.Count -gt 0) {
        # targetId -> dependencyType. Seed with what $To already depends on so
        # existing dependencies survive the replace.
        $desiredDependencies = [ordered]@{}
        foreach ($existing in (Get-IntuneWin32AppDependency -ID $To.id)) {
            if (($existing.PSObject.Properties.Name -contains "targetType") -and ($existing.targetType -eq "parent")) {
                continue
            }
            $desiredDependencies[$existing.targetId] = & $normalizeDependencyType $existing.dependencyType
        }
        foreach ($dependency in $childDependencies) {
            # An app cannot depend on itself, and $To must not inherit a
            # dependency on the app it is replacing.
            if (($dependency.targetId -eq $To.id) -or ($dependency.targetId -eq $From.id)) {
                Write-Log "Skipping child dependency $($dependency.targetId) because it points at the source or target app."
                continue
            }
            $desiredDependencies[$dependency.targetId] = & $normalizeDependencyType $dependency.dependencyType
        }

        if ($desiredDependencies.Count -eq 0) {
            Write-Log "No child dependencies to migrate to $($To.id)."
        }
        else {
            $maxRetries = 3
            $try = 0
            $successfullyAdded = $false
            while (!$successfullyAdded -and ($try++ -lt $maxRetries)) {
                try {
                    $dependencyObjects = @()
                    foreach ($targetId in $desiredDependencies.Keys) {
                        # Returns $null and warns if the target app no longer exists.
                        $dependencyObject = New-IntuneWin32AppDependency -ID $targetId -DependencyType $desiredDependencies[$targetId]
                        if ($dependencyObject) {
                            $dependencyObjects += $dependencyObject
                        }
                        else {
                            Write-Log "Unable to build a dependency object for $targetId. It may no longer exist in Intune."
                        }
                    }
                    if ($dependencyObjects.Count -eq 0) {
                        Write-Log "None of the child dependencies could be resolved. Skipping dependency migration to $($To.id)."
                        break
                    }

                    Add-IntuneWin32AppDependency -ID $To.id -Dependency $dependencyObjects | Out-Null

                    # Add-IntuneWin32AppDependency warns instead of throwing when
                    # Graph rejects the update, so read the result back rather
                    # than assuming the call succeeded.
                    $assignedTargets = @(Get-IntuneWin32AppDependency -ID $To.id |
                        Where-Object { $_.targetType -ne "parent" } |
                        ForEach-Object { $_.targetId })
                    $missingTargets = @($dependencyObjects.targetId | Where-Object { $assignedTargets -notcontains $_ })
                    if ($missingTargets.Count -eq 0) {
                        $successfullyAdded = $true
                        Write-Log "Migrated $($dependencyObjects.Count) child dependency(ies) to $($To.id): $($dependencyObjects.targetId -join ', ')"
                    }
                    else {
                        Write-Log "Dependencies missing from $($To.id) after attempt $($try) of $($maxRetries): $($missingTargets -join ', ')"
                    }
                }
                catch {
                    Write-Log "Failed to migrate child dependencies to $($To.id) on attempt $($try) of $($maxRetries): $_"
                }

                if (!$successfullyAdded -and ($try -lt $maxRetries)) {
                    Start-Sleep -Seconds 2
                }
            }
            if (!$successfullyAdded) {
                Write-Log "Unable to migrate child dependencies to $($To.id) after $maxRetries attempts."
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
        return , @()
    }
    
    # Include the current name, any (N-x) rename, and the {DETECT} anchor for the same app.
    $anchorName = Get-DetectAnchorName -DisplayName $DisplayName
    $sortable = ($AllSimilarApps | Where-Object {
        ($_.DisplayName -eq $DisplayName) -or
        ($_.DisplayName -like "$DisplayName (N-*") -or
        ($_.DisplayName -eq $anchorName)
    })
    # Sort by version descending, then by createdDateTime descending. The unary
    # comma keeps a single match from being unrolled into a scalar, so callers can
    # always index and use .Count safely.
    return , @($sortable | Sort-Object @{Expression = {[VersionPro]$_.displayVersion}; Descending = $true}, @{Expression = "createdDateTime"; Descending = $true})
}



function Invoke-YardstickGraphRequest {
    <#
    .SYNOPSIS
    Minimal Microsoft Graph wrapper for the handful of calls that IntuneWin32App
    does not expose (assignment-level auto-update settings, relationship
    direction). Reuses the authentication header maintained by
    Connect-AutoMSIntuneGraph / Connect-MSIntuneGraph.

    .PARAMETER Resource
    Graph resource path relative to the API version root, e.g.
    "deviceAppManagement/mobileApps/<id>/assignments".

    .PARAMETER Method
    HTTP method. Defaults to Get.

    .PARAMETER Body
    Object to serialize as the JSON request body.

    .PARAMETER ApiVersion
    "beta" (default) or "v1.0". The auto-update assignment settings and the
    per-app relationships collection are only exposed on beta.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Resource,

        [ValidateSet('Get', 'Post', 'Patch', 'Put', 'Delete')]
        [string]$Method = 'Get',

        $Body,

        [ValidateSet('beta', 'v1.0')]
        [string]$ApiVersion = 'beta'
    )

    if (-not $Global:AuthenticationHeader -or -not $Global:AuthenticationHeader.Authorization) {
        throw "Graph authentication header is missing. Call Connect-AutoMSIntuneGraph before using Invoke-YardstickGraphRequest."
    }

    $headers = @{
        Authorization  = $Global:AuthenticationHeader.Authorization
        'Content-Type' = 'application/json'
    }

    $uri = "https://graph.microsoft.com/$ApiVersion/$($Resource.TrimStart('/'))"
    $params = @{
        Uri         = $uri
        Headers     = $headers
        Method      = $Method
        ErrorAction = 'Stop'
    }
    if ($null -ne $Body) {
        $params['Body'] = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 }
    }

    $response = Invoke-RestMethod @params

    # Unwrap OData collections and follow paging so callers always get a flat array.
    if ($null -ne $response -and $response.PSObject.Properties['value']) {
        $results = [System.Collections.Generic.List[object]]::new()
        foreach ($item in $response.value) { $results.Add($item) | Out-Null }
        $next = $response.'@odata.nextLink'
        while ($next) {
            $page = Invoke-RestMethod -Uri $next -Headers $headers -Method Get -ErrorAction Stop
            foreach ($item in $page.value) { $results.Add($item) | Out-Null }
            $next = $page.'@odata.nextLink'
        }
        return $results.ToArray()
    }
    return $response
}


function Test-IsVersionDetection {
    <#
    .SYNOPSIS
    Returns $true when the detection type is one that compares against a version
    number and therefore benefits from a {DETECT} anchor to catch older installs.
    #>
    param(
        [string]$DetectionType,
        [string]$FileDetectionMethod,
        [string]$RegistryDetectionMethod
    )
    switch ($DetectionType) {
        'msi'      { return $true }
        'file'     { return ($FileDetectionMethod -eq 'version') }
        'registry' { return ($RegistryDetectionMethod -eq 'version') }
        default    { return $false }
    }
}


function Get-DetectAnchor {
    <#
    .SYNOPSIS
    Returns the "{DETECT} <DisplayName>" Intune app object if one exists, else $null.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$DisplayName
    )
    $anchorName = Get-DetectAnchorName -DisplayName $DisplayName
    $candidates = Get-IntuneWin32App -DisplayName $anchorName -ErrorAction SilentlyContinue
    if (-not $candidates) { return $null }
    return ($candidates | Where-Object DisplayName -eq $anchorName | Select-Object -First 1)
}


function Get-DetectAnchorName {
    <#
    .SYNOPSIS
    Returns the reserved Intune display name used for an application's detection anchor.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$DisplayName
    )
    return "{DETECT} $DisplayName"
}


function Set-DetectAnchor {
    <#
    .SYNOPSIS
    Renames the given Intune app to "{DETECT} <DisplayName>" so it is preserved
    across retention cleanup as the low-water-mark detection anchor. Idempotent.
    #>
    param(
        [Parameter(Mandatory=$true)]
        $App,
        [Parameter(Mandatory=$true)]
        [string]$DisplayName
    )
    $anchorName = Get-DetectAnchorName -DisplayName $DisplayName
    if ($App.DisplayName -eq $anchorName) {
        Write-Log "App $($App.Id) is already the {DETECT} anchor for $DisplayName"
        return
    }
    Write-Log "Pinning $($App.DisplayName) ($($App.Id)) as {DETECT} anchor for $DisplayName"
    Set-IntuneWin32App -Id $App.Id -DisplayName $anchorName | Out-Null
}


function Get-YardstickSupersedenceRelationship {
    <#
    .SYNOPSIS
    Returns the supersedence relationships that touch the given Win32 app, in both
    directions.

    .DESCRIPTION
    The beta `mobileApps/{id}/relationships` collection contains an entry for every
    supersedence edge the app participates in. Entries where `sourceId` equals the
    app id are forward links (this app supersedes `targetId`); entries where
    `targetId` equals the app id are reverse links (`sourceId` supersedes this app).
    Reverse links must be cleared from the *parent* before Intune will allow the
    child to be deleted.

    .PARAMETER Id
    The Win32 app id to inspect.

    .PARAMETER Direction
    'Forward' (this app supersedes others), 'Reverse' (others supersede this app),
    or 'All' (default).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id,

        [ValidateSet('All', 'Forward', 'Reverse')]
        [string]$Direction = 'All'
    )

    try {
        $relationships = @(Invoke-YardstickGraphRequest -Resource "deviceAppManagement/mobileApps/$Id/relationships" |
            Where-Object { $_.'@odata.type' -eq '#microsoft.graph.mobileAppSupersedence' })
    } catch {
        Write-Log "WARNING: Failed to read supersedence relationships for $Id : $_"
        return ,@()
    }

    # Every return is comma-wrapped: `return @($x)` unrolls a single-element array
    # back to a bare object on the way out, and [PSCustomObject] has no synthetic
    # .Count (unlike other scalars in PS 3.0+), so a caller testing
    # `(Get-YardstickSupersedenceRelationship ...).Count -gt 0` silently sees $null
    # and concludes there are no links. That skipped the stale-link strip in
    # Set-YardstickSupersedence for exactly the one-link case, leaving a target
    # still pointing at the new app - which Intune then rejects with
    # "A circular dependency was created while adding app relationships."
    switch ($Direction) {
        # sourceId is only null on freshly-submitted payloads; Graph backfills it
        # with the parent id, so treat null as "this app is the parent".
        'Forward' { return ,@($relationships | Where-Object { (-not $_.sourceId) -or ($_.sourceId -eq $Id) }) }
        'Reverse' { return ,@($relationships | Where-Object { $_.sourceId -and ($_.sourceId -ne $Id) -and ($_.targetId -eq $Id) }) }
        default   { return ,$relationships }
    }
}


function Remove-SupersedenceReference {
    <#
    .SYNOPSIS
    Removes a single supersedence target from a parent app, leaving the parent's
    other supersedence targets (and all of its dependencies) intact.

    .DESCRIPTION
    Intune has no "delete one relationship" operation for Win32 apps - the whole
    relationship set is replaced via updateRelationships. This rebuilds the parent's
    supersedence list without $TargetId and re-submits it.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ParentId,

        [Parameter(Mandatory = $true)]
        [string]$TargetId
    )

    $forward = @(Get-YardstickSupersedenceRelationship -Id $ParentId -Direction Forward)
    if (-not ($forward | Where-Object targetId -eq $TargetId)) {
        return
    }

    $remaining = @($forward | Where-Object targetId -ne $TargetId)
    if ($remaining.Count -eq 0) {
        Write-Log "Clearing supersedence on $ParentId (was its only target: $TargetId)"
        Remove-IntuneWin32AppSupersedence -ID $ParentId | Out-Null
        return
    }

    $rebuilt = New-SupersedenceObject -Relationships $remaining
    if ($rebuilt.Count -eq 0) {
        Write-Log "WARNING: Could not rebuild supersedence for $ParentId; leaving it untouched"
        return
    }
    Write-Log "Rebuilding supersedence on $ParentId without target $TargetId ($($rebuilt.Count) remaining)"
    Add-IntuneWin32AppSupersedence -ID $ParentId -Supersedence $rebuilt | Out-Null
}


function New-SupersedenceObject {
    <#
    .SYNOPSIS
    Builds the OrderedDictionary array that Add-IntuneWin32AppSupersedence expects.

    .DESCRIPTION
    Accepts either existing relationship objects (which carry targetId +
    supersedenceType) or an explicit target id / type pair. Any entry that cannot
    be resolved is dropped rather than passed through as $null, because
    Add-IntuneWin32AppSupersedence declares [OrderedDictionary[]] with
    ValidateNotNullOrEmpty and would otherwise fail the whole batch.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Relationships')]
    [OutputType([System.Collections.Specialized.OrderedDictionary[]])]
    param(
        [Parameter(ParameterSetName = 'Relationships')]
        [AllowEmptyCollection()]
        [array]$Relationships = @(),

        [Parameter(ParameterSetName = 'Explicit')]
        [AllowEmptyCollection()]
        [array]$TargetIds = @(),

        [Parameter(ParameterSetName = 'Explicit')]
        [ValidateSet('Update', 'Replace')]
        [string]$Type = 'Update'
    )

    $pairs = if ($PSCmdlet.ParameterSetName -eq 'Relationships') {
        foreach ($relationship in $Relationships) {
            # supersedenceType comes back lowercase from Graph; New-IntuneWin32AppSupersedence validates Update/Replace.
            $resolved = if ($relationship.supersedenceType -eq 'replace') { 'Replace' } else { 'Update' }
            [PSCustomObject]@{ TargetId = $relationship.targetId; Type = $resolved }
        }
    } else {
        foreach ($targetId in $TargetIds) {
            [PSCustomObject]@{ TargetId = $targetId; Type = $Type }
        }
    }

    $built = [System.Collections.Generic.List[System.Collections.Specialized.OrderedDictionary]]::new()
    foreach ($pair in $pairs) {
        if (-not $pair.TargetId) { continue }
        $object = New-IntuneWin32AppSupersedence -ID $pair.TargetId -SupersedenceType $pair.Type
        if ($object) {
            $built.Add([System.Collections.Specialized.OrderedDictionary]$object) | Out-Null
        } else {
            Write-Log "WARNING: Could not build supersedence object for target $($pair.TargetId) - skipping"
        }
    }
    return , $built.ToArray()
}


function Set-YardstickSupersedence {
    <#
    .SYNOPSIS
    Makes $NewApp the single superseding parent for every entry in $SupersededApps.

    .DESCRIPTION
    Yardstick maintains a flat supersedence graph: only the newest version ever
    supersedes anything, so stale forward links are stripped off the targets first
    and the parent's own list is rebuilt from scratch. This keeps re-runs idempotent
    and keeps the graph inside Intune's 10-node limit.

    .PARAMETER Type
    "Update" (in-place upgrade) or "Replace" (uninstall previous version first).

    .PARAMETER UpdateOnlyIds
    App ids that must always be superseded with "Update" regardless of $Type. The
    {DETECT} anchor lives here: its detection rule deliberately matches a very wide
    version range, so a "Replace" against it would uninstall the app from every
    device that has any version installed.
    #>
    param(
        [Parameter(Mandatory=$true)]
        $NewApp,
        [Parameter(Mandatory=$true)]
        [AllowEmptyCollection()]
        [array]$SupersededApps,
        [ValidateSet('Update','Replace')]
        [string]$Type = 'Update',
        [AllowEmptyCollection()]
        [string[]]$UpdateOnlyIds = @()
    )

    $targets = @($SupersededApps | Where-Object { $_ -and $_.id -and $_.id -ne $NewApp.id })

    # De-duplicate: the same app can arrive from both the kept list and the anchor.
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $targets = @($targets | Where-Object { $seen.Add($_.id) })

    if ($targets.Count -eq 0) {
        Write-Log "No supersedence targets for $($NewApp.DisplayName) - clearing any stale links"
        if (@(Get-YardstickSupersedenceRelationship -Id $NewApp.id -Direction Forward).Count -gt 0) {
            Remove-IntuneWin32AppSupersedence -ID $NewApp.id | Out-Null
        }
        return 0
    }

    $updateOnly = [System.Collections.Generic.HashSet[string]]::new([string[]]$UpdateOnlyIds, [StringComparer]::OrdinalIgnoreCase)

    # Intune caps a supersedence graph at 10 nodes, one of which is the parent.
    # Update-only targets (the {DETECT} anchor) are reserved first: the anchor is
    # by definition the oldest version, so a plain newest-first trim would drop
    # exactly the target that catches stale installs.
    $maxTargets = 9
    if ($targets.Count -gt $maxTargets) {
        Write-Log "WARNING: $($targets.Count) supersedence targets exceeds Intune's limit of $maxTargets; trimming the oldest"
        $reserved = @($targets | Where-Object { $updateOnly.Contains($_.id) } | Select-Object -First $maxTargets)
        $fill = @($targets |
            Where-Object { -not $updateOnly.Contains($_.id) } |
            Sort-Object @{Expression = {[VersionPro]$_.displayVersion}; Descending = $true} |
            Select-Object -First ($maxTargets - $reserved.Count))
        $targets = @($fill) + @($reserved)
    }

    # Strip stale forward links off every target so the newest app is the only parent.
    foreach ($target in $targets) {
        try {
            if (@(Get-YardstickSupersedenceRelationship -Id $target.id -Direction Forward).Count -gt 0) {
                Write-Log "Clearing stale supersedence on $($target.DisplayName) ($($target.id))"
                Remove-IntuneWin32AppSupersedence -ID $target.id | Out-Null
            }
        } catch {
            Write-Log "WARNING: Failed to clear existing supersedence on $($target.DisplayName) ($($target.id)): $_"
        }
    }

    $supersedence = [System.Collections.Generic.List[System.Collections.Specialized.OrderedDictionary]]::new()
    foreach ($target in $targets) {
        $targetType = if ($updateOnly.Contains($target.id)) { 'Update' } else { $Type }
        $built = New-SupersedenceObject -TargetIds @($target.id) -Type $targetType
        foreach ($object in $built) { $supersedence.Add($object) | Out-Null }
    }

    if ($supersedence.Count -eq 0) {
        Write-Log "WARNING: No resolvable supersedence targets for $($NewApp.DisplayName) - skipping"
        return 0
    }

    Write-Log "Attaching supersedence ($Type) from $($NewApp.DisplayName) to $($supersedence.Count) target(s)"
    # Add-IntuneWin32AppSupersedence replaces the parent's whole supersedence set,
    # so there is no need to clear it first.
    Add-IntuneWin32AppSupersedence -ID $NewApp.id -Supersedence $supersedence.ToArray() | Out-Null

    # The cmdlet downgrades Graph failures to warnings, so read the graph back and
    # report what Intune actually stored rather than what we asked for.
    $attached = @(Get-YardstickSupersedenceRelationship -Id $NewApp.id -Direction Forward)
    if ($attached.Count -ne $supersedence.Count) {
        Write-Log "ERROR: Expected $($supersedence.Count) supersedence target(s) on $($NewApp.DisplayName) but Intune reports $($attached.Count)"
    }
    return $attached.Count
}


function Set-AssignmentAutoUpdate {
    <#
    .SYNOPSIS
    Turns Intune's native auto-update on for a Win32 app's assignments so devices
    running a superseded version are pulled forward without any custom remediation.

    .DESCRIPTION
    Sets `settings.autoUpdateSettings.autoUpdateSupersededAppsState` on each matching
    assignment. Microsoft only honours this for the *available* intent - "the
    supersedence auto-update only applies for available assignments" - so the default
    filter is 'available' and required assignments are left alone.

    Assignments are read straight from Graph rather than via
    Get-IntuneWin32AppAssignment because that cmdlet does not surface the assignment
    id, which is required to PATCH an individual assignment.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$AppId,
        [bool]$Enabled = $true,
        [ValidateSet('', 'required', 'available')]
        [string]$IntentFilter = 'available'
    )

    try {
        $assignments = @(Invoke-YardstickGraphRequest -Resource "deviceAppManagement/mobileApps/$AppId/assignments")
    } catch {
        Write-Log "WARNING: Failed to read assignments for $AppId while setting auto-update: $_"
        return 0
    }

    if ($IntentFilter) {
        $assignments = @($assignments | Where-Object intent -eq $IntentFilter)
    }
    if ($assignments.Count -eq 0) {
        Write-Log "No $(if ($IntentFilter) { "$IntentFilter-intent " })assignments on $AppId to configure auto-update on"
        return 0
    }

    $desiredState = if ($Enabled) { 'enabled' } else { 'notConfigured' }
    $count = 0
    foreach ($assignment in $assignments) {
        if (-not $assignment.id) { continue }

        # Exclusion targets carry no settings object and cannot auto-update.
        if ($assignment.target.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget') { continue }

        if ($assignment.settings.autoUpdateSettings.autoUpdateSupersededAppsState -eq $desiredState) {
            $count++
            continue
        }

        # Re-send the existing settings so notifications, delivery optimization and
        # install-time settings are not reset by the PATCH.
        $settings = [ordered]@{
            '@odata.type'                  = '#microsoft.graph.win32LobAppAssignmentSettings'
            'notifications'                = if ($assignment.settings.notifications) { $assignment.settings.notifications } else { 'showAll' }
            'restartSettings'              = $assignment.settings.restartSettings
            'deliveryOptimizationPriority' = if ($assignment.settings.deliveryOptimizationPriority) { $assignment.settings.deliveryOptimizationPriority } else { 'notConfigured' }
            'installTimeSettings'          = $assignment.settings.installTimeSettings
            'autoUpdateSettings'           = [ordered]@{
                '@odata.type'                  = '#microsoft.graph.win32LobAppAutoUpdateSettings'
                'autoUpdateSupersededAppsState' = $desiredState
            }
        }

        try {
            Invoke-YardstickGraphRequest -Resource "deviceAppManagement/mobileApps/$AppId/assignments/$($assignment.id)" `
                -Method Patch -Body @{ settings = $settings } | Out-Null
            $count++
        } catch {
            Write-Log "WARNING: Failed to set auto-update on assignment $($assignment.id) for app $AppId : $_"
        }
    }
    Write-Log "Auto-update ($desiredState) is set on $count assignment(s) for $AppId"
    return $count
}


function Remove-YardstickApp {
    <#
    .SYNOPSIS
    Safely deletes an Intune Win32 app.

    .DESCRIPTION
    Intune refuses to delete an app that still participates in a supersedence
    relationship, so both directions are unwound first: any parent that supersedes
    this app has the reference rebuilt without it, and the app's own forward links
    are cleared.
    #>
    param(
        [Parameter(Mandatory=$true)]
        $App
    )

    $relationships = @(Get-YardstickSupersedenceRelationship -Id $App.id)

    $parentIds = @($relationships |
        Where-Object { $_.sourceId -and ($_.sourceId -ne $App.id) -and ($_.targetId -eq $App.id) } |
        Select-Object -ExpandProperty sourceId -Unique)
    foreach ($parentId in $parentIds) {
        try {
            Remove-SupersedenceReference -ParentId $parentId -TargetId $App.id
        } catch {
            Write-Log "WARNING: Failed to detach $($App.DisplayName) ($($App.id)) from superseding app $parentId : $_"
        }
    }

    $hasForwardLinks = @($relationships | Where-Object { (-not $_.sourceId) -or ($_.sourceId -eq $App.id) }).Count -gt 0
    if ($hasForwardLinks) {
        try {
            Remove-IntuneWin32AppSupersedence -ID $App.id | Out-Null
        } catch {
            Write-Log "WARNING: Failed to strip supersedence from $($App.DisplayName) ($($App.id)) before delete: $_"
        }
    }

    Write-Log "Removing app $($App.DisplayName) ($($App.id))"
    Remove-IntuneWin32App -Id $App.id

    # Remove-IntuneWin32App downgrades Graph failures (including Intune's refusal
    # to delete an app that is still in a relationship) to a warning, so confirm
    # the app is really gone. Callers rely on this throwing to know a prune failed.
    $stillPresent = $null
    try {
        $stillPresent = Get-IntuneWin32App -Id $App.id -ErrorAction SilentlyContinue
    } catch {
        $stillPresent = $null
    }
    if ($stillPresent) {
        throw "Intune still reports app $($App.DisplayName) ($($App.id)) after the delete request."
    }
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
        'supersedence', 'uninstallPreviousVersion', 'autoUpdateOnAssignment', 'autoUpdate',
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

    .PARAMETER AutoUpdateStatus
    Optional short status string from the auto-update handler (e.g. "published (12 device(s))",
    "no outdated devices", "skipped (no usable installer URL)"). $null when auto-update was disabled.
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

        [string]$Action = "Updated",

        [AllowNull()]
        [string]$AutoUpdateStatus
    )

    if (-not $Script:SuccessfulApplications) {
        $Script:SuccessfulApplications = [System.Collections.Generic.List[PSObject]]::new()
    }

    $appInfo = [PSCustomObject]@{
        ApplicationId    = $ApplicationId
        DisplayName      = $DisplayName
        Version          = $Version
        Action           = $Action
        Dependents       = $Dependents
        AutoUpdateStatus = $AutoUpdateStatus
        Timestamp        = Get-Date
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
                            <th>Auto-Update</th>
                            <th>Time</th>
                        </tr>
"@
            foreach ($app in $Script:SuccessfulApplications) {
                $auStatus = if ($app.PSObject.Properties['AutoUpdateStatus'] -and $app.AutoUpdateStatus) { $app.AutoUpdateStatus } else { '&mdash;' }
                $auColor = switch -Regex ($auStatus) {
                    '^published'   { '#28a745' }
                    '^no outdated' { '#6c757d' }
                    '^skipped'     { '#ffc107' }
                    '^failed'      { '#dc3545' }
                    default        { '#6c757d' }
                }
                $emailBody += @"
                <tr>
                    <td data-label="Application"><strong>$($app.DisplayName)</strong><br><small>ID: $($app.ApplicationId)</small></td>
                    <td data-label="Version">$($app.Version)</td>
                    <td data-label="Action">$($app.Action)</td>
                    <td data-label="Auto-Update"><span style="color: $auColor;">$auStatus</span></td>
                    <td data-label="Time" class="timestamp">$($app.Timestamp.ToString("MM/dd/yyyy HH:mm:ss"))</td>
                </tr>
"@
                # Add dependency information if present
                if ($app.Dependents -and $app.Dependents.Count -gt 0) {
                    $emailBody += @"
                <tr class="dependents-row">
                    <td colspan="5" style="background-color: #f8f9fa; color:#1a1a1a; padding-left: 30px;">
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
