using module .\VersionPro.psm1

# Global because the runner and the recipes it executes call these directly. No
# -Force: when Yardstick.psm1 (or a test) has already loaded them, re-importing
# would load a second copy and throw away the first one's state, including
# YardstickGraph's cached access token. YardstickLog comes first: the watchdog
# and backup threads import this module by path, and get Write-Log through it.
Import-Module $PSScriptRoot\YardstickLog.psm1 -Scope Global
Import-Module $PSScriptRoot\YardstickGraph.psm1 -Scope Global
Import-Module $PSScriptRoot\YardstickIntune.psm1 -Scope Global



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
    $requiredModules = @('powershell-yaml')
    foreach ($mod in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $mod)) {
            $errors.Add("Required PowerShell module '$mod' is not installed. Install with: Install-Module $mod")
        }
    }

    # Optional PowerShell modules (warn only). Pester is left out: recipes never
    # use it, so a missing copy shouldn't be reported on every run.
    foreach ($mod in (Get-YardstickOptionalModule | Where-Object UsedByRecipes)) {
        if (-not $mod.InstalledVersion) {
            $warnings.Add("Optional PowerShell module '$($mod.Name)' is not installed. It is needed for $($mod.Purpose).")
        } elseif (-not $mod.IsCurrent) {
            $warnings.Add("Optional PowerShell module '$($mod.Name)' $($mod.InstalledVersion) is installed, but recipes need $($mod.MinimumVersion) or later. Run .\Install-Yardstick.ps1 -Step Prerequisites to update it.")
        }
    }

    # External tools. A curl on PATH (Windows ships one in System32) serves the
    # download scripts just as well as the bundled copy, so only warn when neither
    # is available. Ask for it by full filename: a bare "curl" is resolved through
    # PATHEXT, and a shell that hands us a truncated one - Git Bash sets it to
    # .CPL - then misses a curl sitting right there on PATH.
    if (-not (Get-Command "curl.exe" -ErrorAction SilentlyContinue)) {
        $curlPath = Join-Path $ToolsPath "curl.exe"
        if (-not (Test-Path $curlPath)) {
            $warnings.Add("curl.exe not found at '$curlPath' or on PATH. Some download scripts may fail.")
        }
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


function Get-YardstickOptionalModule {
    <#
    .SYNOPSIS
    Lists the optional PowerShell modules and whether a usable version is installed.

    .DESCRIPTION
    The browser-automation recipes call Start-SeDriver and Get-SeElement, which
    only exist in Selenium 4, and Selenium 4 is only published as a prerelease.
    A module's Version property drops the prerelease label (every 4.0.0 preview
    reports 4.0.0), so versions are compared as semantic versions built from
    Version and PrivateData.PSData.Prerelease.

    Pester runs the unit tests. Windows ships Pester 3.4, signed by Microsoft
    rather than the Pester project, so Install-Module refuses a newer version
    unless the publisher check is skipped.

    .OUTPUTS
    One object per module: Name, Purpose, UsedByRecipes, MinimumVersion (a
    semantic version string, or $null for any version), Prerelease (install with
    prerelease versions allowed), SkipPublisherCheck, InstalledVersion (the
    newest installed, or $null) and IsCurrent.
    #>
    $definitions = @(
        @{ Name = 'Selenium'; Purpose = 'recipes that drive a browser'; UsedByRecipes = $true; MinimumVersion = '4.0.0-preview3'; Prerelease = $true; SkipPublisherCheck = $false }
        @{ Name = 'TUN.CredentialManager'; Purpose = 'recipes that sign in to a vendor portal'; UsedByRecipes = $true; MinimumVersion = $null; Prerelease = $false; SkipPublisherCheck = $false }
        @{ Name = 'Pester'; Purpose = 'running the unit tests in Tests\Unit'; UsedByRecipes = $false; MinimumVersion = '5.0.0'; Prerelease = $false; SkipPublisherCheck = $true }
    )
    foreach ($definition in $definitions) {
        $installed = @(Get-Module -ListAvailable -Name $definition.Name | ForEach-Object {
            $label = $_.PrivateData.PSData.Prerelease
            $text = "$($_.Version.Major).$($_.Version.Minor).$([Math]::Max($_.Version.Build, 0))"
            if ($label) { $text += "-$label" }
            [System.Management.Automation.SemanticVersion]$text
        } | Sort-Object -Descending)
        $newest = if ($installed.Count -gt 0) { $installed[0] } else { $null }
        [PSCustomObject]@{
            Name               = $definition.Name
            Purpose            = $definition.Purpose
            UsedByRecipes      = $definition.UsedByRecipes
            MinimumVersion     = $definition.MinimumVersion
            Prerelease         = $definition.Prerelease
            SkipPublisherCheck = $definition.SkipPublisherCheck
            InstalledVersion   = if ($newest) { $newest.ToString() } else { $null }
            IsCurrent          = [bool]($newest -and (-not $definition.MinimumVersion -or $newest -ge [System.Management.Automation.SemanticVersion]$definition.MinimumVersion))
        }
    }
}



function Test-YardstickTransientNetworkError {
    <#
    .SYNOPSIS
    Reports whether an error record was caused by a transient network fault.

    .DESCRIPTION
    Vendor endpoints intermittently drop or time out connections, and a single
    refused socket is enough to fail a recipe that would have succeeded a few
    seconds later. This walks the inner-exception chain looking for the fault
    types that are worth replaying - connection timeouts, refused or unreachable
    hosts, and DNS hiccups - so callers can retry only those and let real
    failures (HTTP 404, bad credentials, a thrown recipe error) surface at once.

    .PARAMETER ErrorRecord
    The error record, exception, or anything carrying an Exception property.

    .OUTPUTS
    [bool] True when the error looks transient and retrying is worthwhile.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        $ErrorRecord
    )

    if ($null -eq $ErrorRecord) { return $false }

    $exception = if ($ErrorRecord -is [System.Exception]) {
        $ErrorRecord
    } elseif ($ErrorRecord.PSObject.Properties['Exception']) {
        $ErrorRecord.Exception
    } else {
        $null
    }

    $transientSocketErrors = @(
        [System.Net.Sockets.SocketError]::TimedOut,
        [System.Net.Sockets.SocketError]::ConnectionRefused,
        [System.Net.Sockets.SocketError]::ConnectionReset,
        [System.Net.Sockets.SocketError]::HostUnreachable,
        [System.Net.Sockets.SocketError]::NetworkUnreachable,
        [System.Net.Sockets.SocketError]::TryAgain
    )
    $transientWebStatuses = @(
        [System.Net.WebExceptionStatus]::Timeout,
        [System.Net.WebExceptionStatus]::ConnectFailure,
        [System.Net.WebExceptionStatus]::NameResolutionFailure,
        [System.Net.WebExceptionStatus]::ReceiveFailure,
        [System.Net.WebExceptionStatus]::SendFailure,
        [System.Net.WebExceptionStatus]::KeepAliveFailure,
        [System.Net.WebExceptionStatus]::PipelineFailure
    )

    # Guard against a self-referential or cyclic inner-exception chain.
    $depth = 0
    while ($exception -and $depth -lt 16) {
        if ($exception -is [System.Net.Sockets.SocketException]) {
            if ($transientSocketErrors -contains $exception.SocketErrorCode) { return $true }
        } elseif ($exception -is [System.Net.WebException]) {
            # A WebException with a Response is a real HTTP answer, not a network fault.
            if ($null -eq $exception.Response -and $transientWebStatuses -contains $exception.Status) { return $true }
        } elseif ($exception -is [System.Net.Http.HttpRequestException]) {
            # HttpRequestException without an HTTP status is a transport failure;
            # with one it carries a real server response and should not be replayed.
            $statusProperty = $exception.PSObject.Properties['StatusCode']
            if (-not $statusProperty -or $null -eq $statusProperty.Value) { return $true }
        } elseif ($exception -is [System.TimeoutException]) {
            return $true
        }

        if ([object]::ReferenceEquals($exception, $exception.InnerException)) { break }
        $exception = $exception.InnerException
        $depth++
    }

    return $false
}



function Get-RedirectedUrl {
    <#
    .SYNOPSIS
    Follows HTTP redirects and returns the final URL.

    .DESCRIPTION
    This function follows HTTP redirects to determine the final destination URL
    of a given URL. Useful for handling shortened URLs or redirects.

    Transient network faults are retried with a short backoff, because this runs
    once per recipe against a vendor endpoint and a single refused connection
    would otherwise fail the whole package.

    .PARAMETER URL
    The URL to follow redirects for.

    .PARAMETER MaximumAttempts
    How many times to attempt the request before giving up. Only transient
    network faults are retried; every other failure throws on the first attempt.

    .OUTPUTS
    The final redirected URL as a string.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [String]$URL,

        [ValidateRange(1, 10)]
        [int]$MaximumAttempts = 3
    )

    $userAgent = [Microsoft.PowerShell.Commands.PSUserAgent]::Chrome

    for ($attempt = 1; ; $attempt++) {
        $httpClient = $null
        $Response = $null

        try {
            $httpClient = [System.Net.Http.HttpClient]::new()
            $httpClient.DefaultRequestHeaders.UserAgent.ParseAdd($userAgent)

            # ResponseHeadersRead stops once the final response headers arrive. Without it
            # HttpClient buffers the entire installer body just to read the resolved URI,
            # which blows past the 100 second timeout on large downloads.
            $Response = $httpClient.GetAsync($URL, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
            if ($Response.IsSuccessStatusCode) {
                $RedirectedURL = $Response.RequestMessage.RequestUri.AbsoluteUri
            } else {
                throw "HTTP request for '$URL' failed with status $($Response.StatusCode)."
            }
            return $RedirectedURL
        } catch {
            if ($attempt -lt $MaximumAttempts -and (Test-YardstickTransientNetworkError -ErrorRecord $_)) {
                Write-Log "Transient network error resolving '$URL' (attempt $attempt/${MaximumAttempts}): $($_.Exception.Message)" -Level Warning
                Start-Sleep -Seconds ([math]::Pow(2, $attempt))
                continue
            }
            throw
        } finally {
            if ($Response) {
                $Response.Dispose()
            }
            if ($httpClient) {
                $httpClient.Dispose()
            }
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
            Write-Log "$Label timed out after $TimeoutSeconds seconds." -Level Warning
            break
        }

        try {
            $result = & $ScriptBlock

            if ($VerifyBlock) {
                $verified = & $VerifyBlock
                if (-not $verified) {
                    $lastError = "Verification failed"
                    Write-Log "$Label failed verification (attempt $attempt/$MaxRetries)." -Level Warning
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
            Write-Log "$Label failed (attempt $attempt/${MaxRetries}): $lastError" -Level Warning

            if ($attempt -lt $MaxRetries -and (Get-Date) -lt $deadline) {
                Start-Sleep -Seconds $DelaySeconds
            }
        }
    }

    Write-Log "$Label failed after $MaxRetries attempts: $lastError" -Level Error
    if ($OnFailure) {
        try {
            & $OnFailure
        }
        catch {
            Write-Log "$Label failure handler failed: $_" -Level Error
        }
    }

    return $null
}



function Get-YardstickBackupFileName {
    <#
    .SYNOPSIS
    Builds the file name a .intunewin backup is stored under.

    .DESCRIPTION
    Returns "<AppId>_<Version>_<yyyyMMdd-HHmmss>.intunewin". The timestamp is the
    moment Intune reported the application as published, so the name records when
    the package was actually shipped rather than when it was built.

    Underscores are the field separator, so any underscore inside AppId or Version
    is replaced along with the characters the filesystem rejects. That keeps the
    name parseable by Get-YardstickBackupTimestamp.

    .PARAMETER AppId
    The recipe id the package was built from.

    .PARAMETER Version
    The application version that was uploaded.

    .PARAMETER Timestamp
    The time the application finished publishing in Intune.

    .OUTPUTS
    String file name (no directory component).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$AppId,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Version,

        [Parameter(Mandatory)]
        [datetime]$Timestamp
    )

    if ([string]::IsNullOrWhiteSpace($AppId)) {
        throw "AppId is required to build a backup file name."
    }
    if ([string]::IsNullOrWhiteSpace($Version)) {
        throw "Version is required to build a backup file name."
    }

    $invalid = [System.IO.Path]::GetInvalidFileNameChars() + [char[]]@('_')

    $sanitize = {
        param([string]$Value)
        $builder = [System.Text.StringBuilder]::new()
        foreach ($char in $Value.Trim().ToCharArray()) {
            if ($invalid -contains $char) { [void]$builder.Append('-') } else { [void]$builder.Append($char) }
        }
        # Collapse runs of '-' so "a//b" does not become "a--b"
        ($builder.ToString() -replace '-{2,}', '-').Trim('-')
    }

    $safeId = & $sanitize $AppId
    $safeVersion = & $sanitize $Version
    $stamp = $Timestamp.ToString('yyyyMMdd-HHmmss', [cultureinfo]::InvariantCulture)

    return "${safeId}_${safeVersion}_${stamp}.intunewin"
}



function Get-YardstickBackupTimestamp {
    <#
    .SYNOPSIS
    Extracts the upload timestamp embedded in a backup file name.

    .DESCRIPTION
    Parses the trailing _yyyyMMdd-HHmmss field written by
    Get-YardstickBackupFileName. Returns $null rather than throwing for names that
    do not conform, so retention can fall back to file metadata for anything that
    predates this naming scheme.

    .PARAMETER FileName
    The file name (or full path) to parse.

    .OUTPUTS
    DateTime, or $null when the name does not carry a parseable timestamp.
    #>
    [CmdletBinding()]
    [OutputType([datetime])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$FileName
    )

    if ([string]::IsNullOrWhiteSpace($FileName)) { return $null }

    $leaf = [System.IO.Path]::GetFileName($FileName)
    # -match rather than -notmatch: only -match reliably populates $matches.
    if (-not ($leaf -match '_(\d{8}-\d{6})\.intunewin$')) { return $null }

    # ParseExact is pinned to the invariant culture - the format is fixed, so the
    # machine's regional settings must not change how it reads.
    $parsed = [datetime]::MinValue
    $ok = [datetime]::TryParseExact(
        $matches[1],
        'yyyyMMdd-HHmmss',
        [cultureinfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::None,
        [ref]$parsed
    )

    if ($ok) { return $parsed }
    return $null
}



function Invoke-YardstickBackupCopy {
    <#
    .SYNOPSIS
    Copies a .intunewin file to the backup share and prunes older backups.

    .DESCRIPTION
    The synchronous body of a backup. Runs either on the main thread or inside the
    thread job started by Start-YardstickBackup, so it takes everything by
    parameter, never calls Write-Log (the job runspace has no $LogLocation), and
    never throws - a backup problem must not be able to fail an application run.
    Log lines are returned for the caller to replay on the main thread.

    The copy lands on a .tmp name and is only renamed into place once its length
    has been verified against the source. That rename is the commit point: a
    process killed mid-copy can only ever leave a .tmp behind, which retention
    ignores and later runs sweep up.

    .PARAMETER SourcePath
    Full path to the .intunewin file to back up.

    .PARAMETER BackupRoot
    Root backup folder. A subfolder named for AppId is created beneath it.

    .PARAMETER AppId
    The recipe id, used as the subfolder name.

    .PARAMETER FileName
    Destination file name, from Get-YardstickBackupFileName.

    .PARAMETER VersionsToKeep
    How many .intunewin files to retain in the app's backup folder (default: 3).
    Zero or negative disables pruning.

    .PARAMETER RemoveSource
    Delete SourcePath when finished. Applied whether or not the copy succeeded.

    .PARAMETER StaleTmpHours
    Age at which abandoned .tmp files are cleaned up (default: 24).

    .OUTPUTS
    PSCustomObject with AppId, Success, Status, BackupPath, Removed,
    SourceRemoved and Log.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath,

        [Parameter(Mandatory)]
        [string]$BackupRoot,

        [Parameter(Mandatory)]
        [string]$AppId,

        [Parameter(Mandatory)]
        [string]$FileName,

        [int]$VersionsToKeep = 3,

        [switch]$RemoveSource,

        [int]$StaleTmpHours = 24
    )

    # A thread job runspace does not inherit this from the parent - it defaults to
    # Continue, which would silently swallow the failures we are trying to report.
    $ErrorActionPreference = 'Stop'

    $log = [System.Collections.Generic.List[string]]::new()
    $removed = [System.Collections.Generic.List[string]]::new()
    $success = $false
    $status = "failed: unknown"
    $backupPath = $null
    $sourceRemoved = $false

    # Bail sentinel. `return` inside the try below would exit the whole function
    # and skip building the result object, so steps that give up throw this after
    # setting $status, and the catch recognises it as already-reported.
    $bail = "YardstickBackupBail"

    try {
        if (-not (Test-Path -LiteralPath $SourcePath)) {
            $status = "failed: source missing"
            $log.Add("Skipped the backup of $AppId because the source file no longer exists ($SourcePath).")
            throw $bail
        }

        $appDir = Join-Path $BackupRoot $AppId
        try {
            if (-not (Test-Path -LiteralPath $appDir)) {
                New-Item -ItemType Directory -Path $appDir -Force | Out-Null
            }
        } catch {
            $status = "failed: backup folder unreachable"
            $log.Add("ERROR: Could not create backup folder '$appDir': $_")
            throw $bail
        }

        $target = Join-Path $appDir $FileName
        $tempTarget = "$target.tmp"

        try {
            Copy-Item -LiteralPath $SourcePath -Destination $tempTarget -Force

            $sourceLength = (Get-Item -LiteralPath $SourcePath).Length
            $copiedLength = (Get-Item -LiteralPath $tempTarget).Length
            if ($copiedLength -ne $sourceLength) {
                throw "copied $copiedLength bytes but source is $sourceLength bytes"
            }

            # Commit point - until this succeeds no complete backup exists.
            Move-Item -LiteralPath $tempTarget -Destination $target -Force
            # Move-Item moves *into* a directory that shares the target's name
            # rather than failing, so confirm we actually landed a file.
            if (-not (Test-Path -LiteralPath $target -PathType Leaf)) {
                throw "destination '$target' is not a file after the move"
            }

            $success = $true
            $backupPath = $target
            $log.Add("Backed up $AppId to $target")
        } catch {
            $status = "failed: $($_.Exception.Message)"
            $log.Add("ERROR: Could not back up $AppId to '$target': $_")
            Remove-Item -LiteralPath $tempTarget -Force -ErrorAction SilentlyContinue
            throw $bail
        }

        # Sweep .tmp files abandoned by a run that was killed mid-copy.
        try {
            $tmpCutoff = (Get-Date).ToUniversalTime().AddHours(-$StaleTmpHours)
            Get-ChildItem -LiteralPath $appDir -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -eq '.tmp' -and $_.LastWriteTimeUtc -lt $tmpCutoff } |
                ForEach-Object {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
                    $log.Add("Removed stale partial backup $($_.Name).")
                }
        } catch {
            $log.Add("WARNING: Could not sweep stale partial backups in '$appDir': $_")
        }

        # Prune to the retention count. Filtering on .Extension rather than
        # -Filter "*.intunewin" avoids the FileSystem provider's 8.3 short-name
        # wildcard matching, which can match names it should not.
        if ($VersionsToKeep -gt 0) {
            try {
                $byTimestamp = @{
                    Expression = {
                        $ts = Get-YardstickBackupTimestamp -FileName $_.Name
                        if ($ts) { $ts } else { $_.LastWriteTimeUtc }
                    }
                    Descending = $true
                }
                $byName = @{ Expression = { $_.Name }; Descending = $true }

                $existing = @(
                    Get-ChildItem -LiteralPath $appDir -File -ErrorAction SilentlyContinue |
                        Where-Object { $_.Extension -eq '.intunewin' } |
                        Sort-Object -Property $byTimestamp, $byName
                )
                foreach ($old in ($existing | Select-Object -Skip $VersionsToKeep)) {
                    Remove-Item -LiteralPath $old.FullName -Force
                    $removed.Add($old.FullName)
                    $log.Add("Pruned old backup $($old.Name).")
                }
                $status = "ok ($([Math]::Min($existing.Count, $VersionsToKeep)) kept)"
            } catch {
                # The backup itself landed, so this is not a failure of the backup.
                $status = "ok (prune failed)"
                $log.Add("WARNING: Could not prune old backups in '$appDir': $_")
            }
        } else {
            $status = "ok (retention disabled)"
        }
    } catch {
        if ($_.Exception.Message -ne $bail) {
            $status = "failed: $($_.Exception.Message)"
            $log.Add("ERROR: Unexpected failure backing up ${AppId}: $_")
        }
    } finally {
        if ($RemoveSource) {
            # Removed regardless of outcome: Invoke-Cleanup would delete it at the
            # end of the run anyway, so keeping a failed copy's source only fills
            # up the Published folder.
            Remove-Item -LiteralPath $SourcePath -Force -ErrorAction SilentlyContinue
            $sourceRemoved = -not (Test-Path -LiteralPath $SourcePath)
        }
    }

    return [PSCustomObject]@{
        AppId         = $AppId
        Success       = $success
        Status        = $status
        BackupPath    = $backupPath
        Removed       = [string[]]$removed
        SourceRemoved = $sourceRemoved
        Log           = [string[]]$log
    }
}



function Start-YardstickBackup {
    <#
    .SYNOPSIS
    Starts a background thread that backs up a .intunewin file.

    .DESCRIPTION
    Hands the copy to a thread job so it overlaps the supersedence and assignment
    work that follows an upload, and registers the job so Wait-YardstickBackup can
    drain it before anything deletes the source.

    A thread job runspace inherits nothing from its parent - not the globally
    imported YardstickSupport module, not $LogLocation, not even
    $ErrorActionPreference - so the scriptblock re-imports this module by absolute
    path and receives its arguments explicitly.

    .PARAMETER SourcePath
    Full path to the staged .intunewin file. The thread deletes it when done.

    .PARAMETER BackupRoot
    Root backup folder from the Backup preference.

    .PARAMETER AppId
    The recipe id, used as the backup subfolder name and to correlate the result.

    .PARAMETER FileName
    Destination file name, from Get-YardstickBackupFileName.

    .PARAMETER VersionsToKeep
    How many .intunewin files to retain per app (default: 3).

    .PARAMETER ModulePath
    Path to this module, re-imported inside the thread. Defaults to this module's
    own location.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseUsingScopeModifierInNewRunspaces', '',
        Justification = 'The thread job scriptblock declares its own param() block and is fed by -ArgumentList, which is deliberate: $using: would silently capture whatever happens to be in scope instead.')]
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath,

        [Parameter(Mandatory)]
        [string]$BackupRoot,

        [Parameter(Mandatory)]
        [string]$AppId,

        [Parameter(Mandatory)]
        [string]$FileName,

        [int]$VersionsToKeep = 3,

        [string]$ModulePath = $PSCommandPath
    )

    if (-not $Script:YardstickBackupJobs) {
        $Script:YardstickBackupJobs = [System.Collections.Generic.List[PSObject]]::new()
    }

    # Explicit rather than relying on auto-loading, so a scheduled task running
    # with $PSModuleAutoLoadingPreference = 'None' fails here with a clear message.
    Import-Module Microsoft.PowerShell.ThreadJob -ErrorAction Stop

    $job = Start-ThreadJob -Name "YardstickBackup_$AppId" -ArgumentList $ModulePath, $SourcePath, $BackupRoot, $AppId, $FileName, $VersionsToKeep -ScriptBlock {
        param($ModulePath, $SourcePath, $BackupRoot, $AppId, $FileName, $VersionsToKeep)
        $ErrorActionPreference = 'Stop'
        Import-Module $ModulePath -Force
        Invoke-YardstickBackupCopy -SourcePath $SourcePath -BackupRoot $BackupRoot `
            -AppId $AppId -FileName $FileName -VersionsToKeep $VersionsToKeep -RemoveSource
    }

    $Script:YardstickBackupJobs.Add([PSCustomObject]@{
        Job        = $job
        AppId      = $AppId
        SourcePath = $SourcePath
        Started    = Get-Date
    })
}


#####################################################################
# RUN SAFETY - nothing in Yardstick may wait forever
#####################################################################
# A 2026-09-25 run sat for two hours inside a single recipe's pre-download script,
# because Invoke-WebRequest's default -TimeoutSec is 0, meaning indefinite, and the
# vendor host accepted the TCP connection and then never answered. The functions
# below bound every operation that can block the runner: HTTP calls, BITS
# downloads, and - as a last resort when something blocks in a way none of those
# cover - the run itself.


function Get-YardstickTimeoutPreference {
    <#
    .SYNOPSIS
    Reads an optional integer timeout out of preferences.yaml.

    .DESCRIPTION
    ConvertFrom-Yaml yields $null for an absent key and "" for a present-but-empty
    one; both mean "use the default". `-as [int]` cannot tell either of those from a
    deliberate 0, which is why this does not just use it.

    An explicit 0 is honoured only when -AllowZero is passed, i.e. only where 0 has
    a defined meaning (disable the bound). Everywhere else a 0 falls back to the
    default, so a stray blank key cannot silently remove a timeout.

    .PARAMETER Value
    The raw value from the preferences hashtable.

    .PARAMETER Default
    The value to use when the key is absent, blank, or unusable.

    .PARAMETER Name
    The preference key name, used only in the warning text.

    .PARAMETER AllowZero
    Treat an explicit 0 as meaningful rather than falling back to -Default.

    .OUTPUTS
    Int32.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        $Value,

        [Parameter(Mandatory)]
        [int]$Default,

        [string]$Name = 'timeout',

        [switch]$AllowZero
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $Default }

    $parsed = $Value -as [int]
    if ($null -eq $parsed -or $parsed -lt 0) {
        Write-Log "Preference '$Name' is set to '$Value', which is not a valid number of units. Using $Default instead." -Level Warning
        return $Default
    }

    if ($parsed -eq 0 -and -not $AllowZero) { return $Default }

    return $parsed
}


function Set-YardstickWebRequestTimeout {
    <#
    .SYNOPSIS
    Applies a default -TimeoutSec to every Invoke-WebRequest/Invoke-RestMethod call.

    .DESCRIPTION
    Recipes run through Invoke-Command -NoNewScope, so a $Global:PSDefaultParameterValues
    entry reaches them: of the 161 recipes that make web calls, only two pass their
    own -TimeoutSec, and patching the rest by hand would not stop the next one being
    written without it. A recipe that does pass -TimeoutSec still wins, because an
    explicit argument always beats a default.

    This bounds the wait for response *headers*, not the response body - verified
    against a server that sent headers immediately then dribbled a body for twelve
    seconds, which -TimeoutSec 4 happily accepted. So this cannot truncate a large
    but healthy -OutFile download; it only ends the case that actually hangs, where
    the far end never answers at all.

    Module code does NOT see this. Each module has its own session state with its
    own $PSDefaultParameterValues, so YardstickGraph and YardstickIntune pass
    -TimeoutSec explicitly instead.

    Call this again at the top of each recipe iteration: a recipe is free to
    reassign $PSDefaultParameterValues, and without re-asserting, one careless
    recipe would un-bound every recipe after it.

    .PARAMETER TimeoutSeconds
    Seconds to wait for response headers. 0 or less removes the defaults, restoring
    PowerShell's indefinite behaviour.
    #>
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [Parameter(Mandatory)]
        [int]$TimeoutSeconds
    )

    if ($null -eq $Global:PSDefaultParameterValues) {
        $Global:PSDefaultParameterValues = @{}
    }

    foreach ($key in 'Invoke-WebRequest:TimeoutSec', 'Invoke-RestMethod:TimeoutSec') {
        if ($TimeoutSeconds -le 0) {
            $Global:PSDefaultParameterValues.Remove($key)
        } else {
            $Global:PSDefaultParameterValues[$key] = $TimeoutSeconds
        }
    }
}


function Invoke-YardstickBitsDownload {
    <#
    .SYNOPSIS
    Downloads a file over BITS with a stall timeout.

    .DESCRIPTION
    Start-BitsTransfer in its default synchronous form blocks the calling thread
    with no way to give up, which makes it one of the few places the runner can
    hang outright. This runs the same transfer asynchronously and watches
    BytesTransferred instead.

    The bound is a *stall* timeout, not a total one: the transfer is abandoned only
    if it moves zero bytes for the whole window. A 2 GB payload on a slow link is
    never punished for being slow, while a connection that dies mid-transfer - or
    one that never gets past Connecting - is cut loose.

    .PARAMETER Source
    URL to download.

    .PARAMETER Destination
    Full path to write to.

    .PARAMETER StallTimeoutMinutes
    Minutes of zero progress before the transfer is abandoned. 0 disables the bound.

    .PARAMETER PollSeconds
    How often to sample BytesTransferred.

    .OUTPUTS
    None. Throws if the transfer fails or stalls.
    #>
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [Parameter(Mandatory)]
        [string]$Source,

        [Parameter(Mandatory)]
        [string]$Destination,

        [int]$StallTimeoutMinutes = 15,

        [ValidateRange(1, 300)]
        [int]$PollSeconds = 5
    )

    $job = Start-BitsTransfer -Source $Source -Destination $Destination -Asynchronous -ErrorAction Stop
    try {
        $lastBytes = [int64]-1
        $lastProgress = Get-Date

        while ($true) {
            $state = [string]$job.JobState

            if ($state -eq 'Transferred') { break }
            if ($state -eq 'Error') {
                throw "BITS transfer of '$Source' failed: $($job.ErrorDescription)"
            }

            $bytes = [int64]$job.BytesTransferred
            if ($bytes -gt $lastBytes) {
                $lastBytes = $bytes
                $lastProgress = Get-Date
                # A multi-gigabyte download can outlast the watchdog's whole stage
                # budget while working perfectly. Report the progress so the budget
                # means "stalled", not "slow".
                Update-YardstickWatchdogProgress
            } elseif ($StallTimeoutMinutes -gt 0 -and
                      ((Get-Date) - $lastProgress).TotalMinutes -ge $StallTimeoutMinutes) {
                throw ("BITS transfer of '$Source' moved no data for $StallTimeoutMinutes minute(s) " +
                       "(state '$state', $bytes of $([int64]$job.BytesTotal) bytes). Abandoning the download.")
            }

            Start-Sleep -Seconds $PollSeconds
        }

        # Only this moves the payload out of its temporary BITS file and into
        # $Destination, so it has to happen before the finally block lets go.
        Complete-BitsTransfer -BitsJob $job
        $job = $null
    } finally {
        # Reached on stall, on error, and on Ctrl-C. A BITS job left behind keeps
        # retrying in the background for days under its own service.
        if ($job) { Remove-BitsTransfer -BitsJob $job -ErrorAction SilentlyContinue }
    }
}


function Test-YardstickWatchdogExpired {
    <#
    .SYNOPSIS
    Decides whether the current recipe stage has overrun its budget.

    .DESCRIPTION
    Split out from the watchdog thread so the decision can be tested without
    threads, clocks, or killing a process.

    .PARAMETER State
    The shared watchdog state, or a snapshot of it. Only StageStarted is read.

    .PARAMETER Now
    The current time, passed in so tests can control it.

    .PARAMETER TimeoutMinutes
    The budget. 0 or less means the watchdog is disabled and this always returns
    false.

    .OUTPUTS
    Boolean.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [hashtable]$State,

        [Parameter(Mandatory)]
        [datetime]$Now,

        [Parameter(Mandatory)]
        [int]$TimeoutMinutes
    )

    if ($TimeoutMinutes -le 0) { return $false }

    $started = $State['StageStarted'] -as [datetime]
    if ($null -eq $started) { return $false }

    return (($Now - $started).TotalMinutes -ge $TimeoutMinutes)
}


function Start-YardstickWatchdog {
    <#
    .SYNOPSIS
    Starts the thread that kills the run if a recipe stage stops making progress.

    .DESCRIPTION
    The backstop for anything the per-operation timeouts do not cover - a native
    call that blocks, a Selenium session that never returns, an unbounded
    Start-BitsTransfer inside a recipe.

    This has to be a thread job rather than a timer. Register-ObjectEvent handlers
    run on the main runspace's event queue, which does not drain while the main
    thread is blocked - exactly the situation the watchdog exists for. A thread job
    runs on a real separate thread in the same process, so it keeps ticking, and
    because it shares the process it can end it.

    As with Start-YardstickBackup, the thread runspace inherits nothing from its
    parent, so the scriptblock re-imports this module by absolute path and takes
    everything else through -ArgumentList. The state hashtable is passed by
    reference and is genuinely shared: a thread job lives in the same AppDomain.

    On expiry the thread logs which application and stage is stuck and for how long,
    then calls [Environment]::Exit so the run ends with a nonzero code the scheduler
    can see. Write-Log closes the file after every line, and the main thread is by
    definition not logging when this fires, so the diagnosis lands on disk.

    With -NoExit the thread only logs the diagnosis, once per stuck stage. That is
    for runs inside someone's interactive session (Invoke-Yardstick at a prompt),
    where ending the process would close their console along with the run.

    .PARAMETER TimeoutMinutes
    Minutes a single stage may go without progress. 0 or less disables the watchdog.

    .PARAMETER PollSeconds
    How often the thread checks.

    .PARAMETER ExitCode
    Process exit code used when the watchdog fires.

    .PARAMETER LogLocation
    Folder holding the log; the thread needs its own copy for Write-Log.

    .PARAMETER LogFile
    Log file name; the thread needs its own copy for Write-Log.

    .PARAMETER ModulePath
    Path to this module, re-imported inside the thread.

    .PARAMETER NoExit
    Log the diagnosis but leave the process running. See the description.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseUsingScopeModifierInNewRunspaces', '',
        Justification = 'The thread job scriptblock declares its own param() block and is fed by -ArgumentList, which is deliberate: $using: would silently capture whatever happens to be in scope instead.')]
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [Parameter(Mandatory)]
        [int]$TimeoutMinutes,

        [ValidateRange(1, 3600)]
        [int]$PollSeconds = 30,

        [int]$ExitCode = 2,

        [string]$LogLocation = $Global:LogLocation,

        [string]$LogFile = $Global:LogFile,

        [string]$ModulePath = $PSCommandPath,

        [switch]$NoExit
    )

    if ($TimeoutMinutes -le 0) {
        Write-Log "Stage watchdog is disabled (recipeStageTimeoutMinutes is 0)."
        return
    }

    Stop-YardstickWatchdog

    # Explicit rather than relying on auto-loading, so a scheduled task running
    # with $PSModuleAutoLoadingPreference = 'None' fails here with a clear message.
    Import-Module Microsoft.PowerShell.ThreadJob -ErrorAction Stop

    $state = [hashtable]::Synchronized(@{
        AppId        = '(startup)'
        Stage        = 'Starting up'
        StageStarted = Get-Date
        WarnedFor    = $null
        ExpiredFor   = $null
        Stop         = $false
    })
    $Script:YardstickWatchdogState = $state

    $Script:YardstickWatchdogJob = Start-ThreadJob -Name 'YardstickWatchdog' `
        -ArgumentList $ModulePath, $state, $TimeoutMinutes, $PollSeconds, $ExitCode, $LogLocation, $LogFile, $NoExit.IsPresent `
        -ScriptBlock {
            param($ModulePath, $State, $TimeoutMinutes, $PollSeconds, $ExitCode, $LogLocation, $LogFile, $NoExit)
            $ErrorActionPreference = 'Stop'
            Import-Module $ModulePath -Force
            $Global:LogLocation = $LogLocation
            $Global:LogFile = $LogFile

            while (-not $State['Stop']) {
                Start-Sleep -Seconds $PollSeconds
                if ($State['Stop']) { break }

                # Snapshot first: the main thread can move to the next stage between
                # reads, and reporting one app's name against another's clock would
                # send whoever reads the log after the wrong recipe.
                $stage = [string]$State['Stage']
                $appId = [string]$State['AppId']
                $started = $State['StageStarted']
                $snapshot = @{ StageStarted = $started }
                $now = Get-Date

                if ((Test-YardstickWatchdogExpired -State $snapshot -Now $now -TimeoutMinutes $TimeoutMinutes) -and
                    ($State['ExpiredFor'] -ne $started)) {
                    $State['ExpiredFor'] = $started
                    $minutes = [math]::Round(($now - $started).TotalMinutes, 1)
                    $outcome = if ($NoExit) { 'Not ending the run, because it shares this PowerShell session.' } else { 'Ending the run.' }
                    # One multi-line entry: the log report folds the indented lines into it.
                    $banner = @(
                        "WATCHDOG: no progress for $minutes minutes (limit $TimeoutMinutes). $outcome"
                        "  '$appId' is stuck in stage '$stage', which began at $(Get-Date $started -Format 'MM/dd/yyyy HH:mm:ss')."
                        "  Raise recipeStageTimeoutMinutes in preferences.yaml if this app is legitimately slow."
                    )
                    if ($NoExit) { $banner += "  Press Ctrl+C to stop the run, or close the window if that does not respond." }
                    $level = if ($NoExit) { 'Warning' } else { 'Error' }
                    Write-Log ($banner -join [Environment]::NewLine) -Level $level -NoContext
                    if (-not $NoExit) {
                        Start-Sleep -Milliseconds 500
                        [Environment]::Exit($ExitCode)
                    }
                }

                # One warning per stage, halfway through the budget, so a run that is
                # merely slow leaves a trail before anything is killed.
                if ((($now - $started).TotalMinutes -ge ($TimeoutMinutes / 2)) -and
                    ($State['WarnedFor'] -ne $started)) {
                    $State['WarnedFor'] = $started
                    $minutes = [math]::Round(($now - $started).TotalMinutes, 1)
                    Write-Log "'$appId' has been in stage '$stage' for $minutes minutes (watchdog limit $TimeoutMinutes)." -Level Warning -NoContext
                }
            }
        }

    $consequence = if ($NoExit) { 'it is logged as stuck' } else { 'the run is ended' }
    Write-Log "Stage watchdog started: a recipe stage may run $TimeoutMinutes minutes without progress before $consequence."
}


function Set-YardstickWatchdogStage {
    <#
    .SYNOPSIS
    Tells the watchdog which application and stage the runner is on, and resets its
    clock.

    .DESCRIPTION
    The budget is per stage rather than per recipe. That is deliberately more
    forgiving - a genuinely long recipe that keeps moving through stages never
    trips it - and it is what lets the diagnosis name the exact point of the stall
    instead of just the application.

    Safe to call when no watchdog is running; it simply does nothing.

    .PARAMETER AppId
    Recipe id currently being processed. Omit to keep the current one.

    .PARAMETER Stage
    Human-readable stage name, quoted verbatim into the watchdog's log output.
    #>
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [string]$AppId,

        [Parameter(Mandatory)]
        [string]$Stage
    )

    $state = $Script:YardstickWatchdogState
    if (-not $state) { return }

    if (-not [string]::IsNullOrWhiteSpace($AppId)) { $state['AppId'] = $AppId }
    $state['Stage'] = $Stage
    $state['StageStarted'] = Get-Date
}


function Update-YardstickWatchdogProgress {
    <#
    .SYNOPSIS
    Tells the watchdog the current stage is still moving, without changing which
    stage it is.

    .DESCRIPTION
    Two stages can legitimately outlast the whole budget on one pass: downloading
    and uploading a large package. clion was 1.6 GB last run, and at 60 minutes a
    link slower than about 0.5 Mbit/s would have the watchdog kill a download that
    was working perfectly. That is a worse failure than the hang it exists to
    prevent, because it would hit good runs.

    So the long-running loops call this every time bytes actually move. The budget
    then means "no progress for N minutes" rather than "no stage change for N
    minutes", which is the thing actually worth killing a run over.

    Callers outside YardstickSupport must guard with Get-Command - each module has
    its own session state, and the unit tests import modules individually.

    Safe to call when no watchdog is running; it simply does nothing.
    #>
    [CmdletBinding()]
    [OutputType([void])]
    param()

    $state = $Script:YardstickWatchdogState
    if (-not $state) { return }

    # Resetting StageStarted also re-arms the half-budget warning, which is keyed
    # on this timestamp.
    $state['StageStarted'] = Get-Date
}


function Get-YardstickWatchdogState {
    <#
    .SYNOPSIS
    Returns the shared watchdog state, or $null when no watchdog is running.

    .DESCRIPTION
    The same object the watchdog thread reads. Exposed for tests and for anything
    that wants to know what stage the runner believes it is on.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    return $Script:YardstickWatchdogState
}


function Stop-YardstickWatchdog {
    <#
    .SYNOPSIS
    Stops the watchdog thread at the end of a run.

    .DESCRIPTION
    Signals the thread to stop and cleans up the job. Does not wait a full poll
    interval for the thread to notice - the process is ending anyway, and a run
    should not spend 30 seconds on teardown.

    Safe to call when no watchdog is running.
    #>
    [CmdletBinding()]
    [OutputType([void])]
    param()

    if ($Script:YardstickWatchdogState) {
        $Script:YardstickWatchdogState['Stop'] = $true
    }

    if ($Script:YardstickWatchdogJob) {
        $null = Wait-Job -Job $Script:YardstickWatchdogJob -Timeout 2 -ErrorAction SilentlyContinue
        Stop-Job -Job $Script:YardstickWatchdogJob -ErrorAction SilentlyContinue
        Remove-Job -Job $Script:YardstickWatchdogJob -Force -ErrorAction SilentlyContinue
    }

    $Script:YardstickWatchdogState = $null
    $Script:YardstickWatchdogJob = $null
}


function Test-YardstickDedicatedProcess {
    <#
    .SYNOPSIS
    Reports whether this PowerShell process was started just to run a script or
    command, so ending it ends nothing but the run.

    .DESCRIPTION
    The watchdog can only stop a stuck run by ending the process. That is right for
    the scheduled task (pwsh -File Yardstick.ps1 ...) and wrong at an interactive
    prompt, where it would close the user's console too.

    Only the pwsh switches before -File/-Command are considered, since everything
    after belongs to the script. -NoExit there means the console outlives the
    script, so the process is not dedicated. When unsure (pwsh started with a bare
    script path, say) this answers false: a stuck run left waiting for Ctrl+C is
    better than a console closed without warning.

    .PARAMETER CommandLineArgs
    The process's arguments, executable first. Defaults to this process's.

    .OUTPUTS
    Boolean.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [string[]]$CommandLineArgs = [Environment]::GetCommandLineArgs()
    )

    $noExit = $false
    foreach ($argument in @($CommandLineArgs | Select-Object -Skip 1)) {
        if ($argument -notmatch '^(-{1,2}|/)(?<name>\w+)') { continue }
        $name = $Matches['name'].ToLowerInvariant()
        if ($name -in 'noe', 'noex', 'noexi', 'noexit') { $noExit = $true; continue }
        # pwsh accepts any unambiguous prefix; -e and -ec are documented aliases.
        if ('file'.StartsWith($name) -or 'command'.StartsWith($name) -or
            'encodedcommand'.StartsWith($name) -or $name -eq 'ec') {
            return (-not $noExit)
        }
    }

    return $false
}



function Get-YardstickBackupInFlight {
    <#
    .SYNOPSIS
    Returns the source paths of backups that are still being copied.

    .DESCRIPTION
    Used to confirm that nothing is mid-copy before a caller deletes files from
    the Published folder.

    .OUTPUTS
    String array of source paths. Empty when nothing is in flight.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param()

    if (-not $Script:YardstickBackupJobs) { return [string[]]@() }

    return [string[]]@(
        $Script:YardstickBackupJobs |
            Where-Object { $_.Job.State -eq 'Running' -or $_.Job.State -eq 'NotStarted' } |
            ForEach-Object { $_.SourcePath }
    )
}



function Wait-YardstickBackup {
    <#
    .SYNOPSIS
    Waits for outstanding backup threads and reports their results.

    .DESCRIPTION
    Drains every job registered by Start-YardstickBackup, replays the log lines
    each one collected (Write-Log is not thread-safe and the job runspace has no
    log configuration, so logging is deferred to here), and records the outcome on
    the matching successful-application entry for the email report.

    Safe to call when nothing is registered, which it is - the script drains at
    several points to guarantee no other code deletes a .intunewin mid-copy.

    .PARAMETER TimeoutSeconds
    Total time to wait for all outstanding jobs (default: 600). Jobs still running
    when it expires are stopped and reported as timed out.

    .OUTPUTS
    The result objects returned by Invoke-YardstickBackupCopy.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject[]])]
    param(
        [int]$TimeoutSeconds = 600
    )

    if (-not $Script:YardstickBackupJobs -or $Script:YardstickBackupJobs.Count -eq 0) {
        return @()
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $results = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($entry in @($Script:YardstickBackupJobs)) {
        $result = $null
        try {
            $remaining = [int][Math]::Max(0, ($deadline - (Get-Date)).TotalSeconds)
            $finished = Wait-Job -Job $entry.Job -Timeout $remaining

            if (-not $finished) {
                Stop-Job -Job $entry.Job -ErrorAction SilentlyContinue
                Write-Log "Backup of $($entry.AppId) timed out after $TimeoutSeconds seconds and was stopped." -Level Warning
                $result = [PSCustomObject]@{
                    AppId = $entry.AppId; Success = $false; Status = "timed out"
                    BackupPath = $null; Removed = [string[]]@(); SourceRemoved = $false
                    Log = [string[]]@()
                }
            } else {
                # Same process, so results come back as live objects with no
                # serialization loss.
                $result = Receive-Job -Job $entry.Job -ErrorAction Stop | Select-Object -Last 1
            }
        } catch {
            Write-Log "Backup of $($entry.AppId) failed: $_" -Level Warning
            $result = [PSCustomObject]@{
                AppId = $entry.AppId; Success = $false; Status = "failed: $_"
                BackupPath = $null; Removed = [string[]]@(); SourceRemoved = $false
                Log = [string[]]@()
            }
        } finally {
            Remove-Job -Job $entry.Job -Force -ErrorAction SilentlyContinue
        }

        if (-not $result) {
            $result = [PSCustomObject]@{
                AppId = $entry.AppId; Success = $false; Status = "failed: no result returned"
                BackupPath = $null; Removed = [string[]]@(); SourceRemoved = $false
                Log = [string[]]@()
            }
        }

        # The job cannot call Write-Log, so it returns its lines with the old text
        # prefixes. Turn those back into levels; routine lines are Detail.
        foreach ($line in @($result.Log)) {
            switch -Regex ($line) {
                '^ERROR: ([\s\S]*)$'   { Write-Log $Matches[1] -Level Error; continue }
                '^WARNING: ([\s\S]*)$' { Write-Log $Matches[1] -Level Warning; continue }
                default                { Write-Log $line -Level Detail }
            }
        }

        # Record on the email report entry for this app, if it has one.
        if ($Script:SuccessfulApplications) {
            $appEntry = $Script:SuccessfulApplications | Where-Object ApplicationId -eq $result.AppId | Select-Object -Last 1
            if ($appEntry -and $appEntry.PSObject.Properties['BackupStatus']) {
                $appEntry.BackupStatus = $result.Status
            }
        }

        $results.Add($result)
    }

    $Script:YardstickBackupJobs.Clear()
    return $results.ToArray()
}



function Get-YardstickAppAssignment {
    <#
    .SYNOPSIS
    Returns the native Yardstick assignment projection as an array.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id
    )

    return @(Get-YardstickWin32AppAssignment -Id $Id | Where-Object { $_.Type -or $_.GroupID -or $_.Intent })
}


function Test-YardstickAssignmentPresent {
    <#
    .SYNOPSIS
    Returns $true when the given app already carries an assignment for a target.

    .DESCRIPTION
    The Add-YardstickWin32AppAssignment* cmdlets downgrade Graph failures and
    duplicate-target conflicts to warnings and emit nothing on the success
    stream, so the fact that one of them returned proves nothing. Reading the
    assignment back off the target app is the only authoritative signal, and it
    doubles as the idempotency check: Intune rejecting an add with "The MobileApp
    Assignment already exists" means the assignment we wanted is there, which is
    a success rather than something to retry.

    Reads through Graph directly rather than Get-YardstickWin32AppAssignment. Under
    Windows PowerShell 5.1 that cmdlet returns $null for an app holding exactly
    one assignment (see the note in Move-AssignmentsAndDependencies), which would
    turn this check into a false negative - and a false negative here is what
    decides whether a source assignment is preserved or deleted.

    .PARAMETER CountKey
    The same key used for removal - GroupID when there is one, otherwise the
    target type - so virtual targets (All Devices/All Users, which carry no
    GroupID) compare by target type instead of collapsing onto a null GroupID.

    .PARAMETER TargetType
    The assignment's '@odata.type'. Distinguishes an include from an exclusion
    for the same group, which CountKey alone cannot. Ignored when either side
    does not report it, so a partial record cannot veto an otherwise clear match.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppId,

        [Parameter(Mandatory = $true)]
        [string]$CountKey,

        [string]$Intent,

        [string]$TargetType
    )

    try {
        $assignments = @(Invoke-YardstickGraphRequest -Resource "deviceAppManagement/mobileApps/$AppId/assignments")
    } catch {
        Write-Log "Could not read the assignments on $AppId to verify ${CountKey}: $_" -Level Warning
        return $false
    }

    foreach ($assignment in $assignments) {
        $existingType = $assignment.target.'@odata.type'
        $existingKey = if ($assignment.target.groupId) { $assignment.target.groupId } else { $existingType }
        if ($existingKey -ne $CountKey) { continue }
        if ($Intent -and ($assignment.intent -ne $Intent)) { continue }
        if ($TargetType -and $existingType -and ($existingType -ne $TargetType)) { continue }
        return $true
    }
    return $false
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
    Hashtable of options for dependent link updates (Enabled, RetryCount, RetryDelaySeconds, TimeoutSeconds, Blocklist).

    .PARAMETER DependentUpdateStatus
    IDictionary to record the status of dependent application updates.

    .PARAMETER ProtectedSourceIds
    HashSet to record source application IDs that should be protected from deletion.

    .PARAMETER ExcludeDependencyTargetIds
    Application IDs that must not be added as dependencies of $To. Callers pass the
    apps they intend to delete this run: copying a dependency that points at one of
    them would create a fresh link that then blocks its deletion.

    .PARAMETER IntentFilter
    When set, only assignments matching this intent are processed (e.g. 'required',
    'available'). Empty (default) processes all intents.

    .PARAMETER CopyOnly
    When set, assignments are added to $To but not removed from $From.

    .PARAMETER SkipDependencies
    When set, dependency migration (child + parent link rewriting) is skipped.
    Useful when calling this function twice for different intents against the
    same From/To pair - dependencies only need to be migrated once.

    .PARAMETER RetryDelaySeconds
    Seconds to wait between assignment add/remove attempts (default: 2). Exists
    so tests can drive the retry paths without sleeping.

    .NOTES
    All Devices and All Users assignments that use an assignment filter are not
    migrated. Add-YardstickWin32AppAssignmentAllDevices/AllUsers accept a filter by
    name only, while Get-YardstickWin32AppAssignment reports the filter id, so the
    filter cannot be carried across - and migrating without it would widen the
    assignment to every device or user. Those assignments are logged and left on
    $From, and $From is added to $ProtectedSourceIds so retention will not delete
    it: deleting the app would destroy targeting nobody can recreate. The same
    protection applies when an assignment could not be added to $To at all. Such an
    app is superseded instead of pruned, and an admin can migrate the assignment by
    hand.

    Removal matches on the target alone (a group id, or the virtual target type)
    because the Yardstick Intune module cannot delete an individual assignment. If
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
        [string[]] $ExcludeDependencyTargetIds = @(),
        [ValidateSet('', 'required', 'available')]
        [string] $IntentFilter = '',
        [switch] $CopyOnly,
        [switch] $SkipDependencies,
        [int] $RetryDelaySeconds = 2
    )

    # If IDs or Display Names were provided instead of application objects, get the app objects
    if ($From -is [String]) {
        if ($From -match "^[0-9a-fA-F\-]{36}$") {
            # Looks like a GUID
            $From = Get-YardstickWin32App -Id $From
        }
        else {
            # Only return exact matches
            $Apps = Get-YardstickWin32App -DisplayName $From | Where-Object DisplayName -eq $From
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
            $To = Get-YardstickWin32App -Id $To
        }
        else {
            # Only return exact matches
            $Apps = Get-YardstickWin32App -DisplayName $To | Where-Object DisplayName -eq $To
            if ($Apps.Count -eq 1) {
                $To = $Apps[0]
            } elseif ($Apps.Count -gt 1) {
                throw "Multiple applications found with the display name '$To'. Please specify by ID instead."
            } else {
                throw "No application found with the display name '$To'."
            }
        }
    }
    Write-Log "Moving assignments and dependencies from $($From.id) to $($To.id)..."
    # The native projection preserves zero, one, and many assignment results.
    $FromAssignments = Get-YardstickAppAssignment -Id $From.id
    $FromDependencies = Get-YardstickWin32AppDependency -Id $From.id
    # Kept as DateTime, not a formatted string: rebuilding a date by formatting
    # to "MM/dd/yyyy" and parsing it back with Get-Date makes the result depend
    # on the host's culture, so on a dd/MM/yyyy host 08/04 silently becomes
    # 8 April and a day past the 12th throws outright.
    #
    # Kind Unspecified, not Local: these dates carry an install time's wall
    # clock, and a Local kind invites a later ToUniversalTime() to shift it by
    # the build host's offset.
    $AvailableDate = [datetime]::SpecifyKind((Get-Date).AddDays($AvailableDateOffset).Date, [System.DateTimeKind]::Unspecified)
    $DeadlineDate = [datetime]::SpecifyKind((Get-Date).AddDays($DeadlineDateOffset).Date, [System.DateTimeKind]::Unspecified)
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
        Blocklist = @()
    }
    if ($DependentLinkOptions) {
        foreach ($key in $DependentLinkOptions.Keys) {
            if ($null -ne $DependentLinkOptions[$key]) {
                $resolvedDependentOptions[$key] = $DependentLinkOptions[$key]
            }
        }
    }
    if (-not ($resolvedDependentOptions.Blocklist -is [System.Collections.IEnumerable])) {
        $resolvedDependentOptions.Blocklist = @($resolvedDependentOptions.Blocklist)
    }
    $normalizedBlocklist = @()
    if ($resolvedDependentOptions.Blocklist) {
        $normalizedBlocklist = @($resolvedDependentOptions.Blocklist | ForEach-Object { $_.ToString().ToLowerInvariant() })
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
                Write-Log "Skipping an assignment with no group ID and unsupported target type '$targetType'. $($From.id) will not be deleted while it is there." -Level Warning
                & $addProtectedSourceId $From.id
                continue
            }
            if ($IntentFilter -and ($Assignment.Intent -ne $IntentFilter)) {
                Write-Log "Skipping the $assignmentLabel assignment: intent $($Assignment.Intent) does not match filter $IntentFilter." -Level Detail
                continue
            }
            # Add-YardstickWin32AppAssignmentAllDevices/AllUsers take a filter by
            # name only, and Get-YardstickWin32AppAssignment reports the filter id,
            # so the filter cannot be carried across. Dropping it would widen the
            # assignment to every device or user, so refuse to migrate instead -
            # and protect $From from deletion, because pruning it would destroy
            # targeting that cannot be recreated from what Intune reports.
            if ($isVirtualTarget -and $hasFilter) {
                Write-Log -Level Warning "Skipping $assignmentLabel assignment because filter $($Assignment.FilterID) ($($Assignment.FilterType)) cannot be reapplied by ID. Migrate this assignment manually; $($From.id) will not be deleted while it is there."
                & $addProtectedSourceId $From.id
                continue
            }

            $maxRetries = 3
            $try = 0
            $successfullyAdded = $false
            Write-Verbose $Assignment
            Write-Log "Processing $assignmentLabel assignment with intent $($Assignment.Intent)$(if ($isExclusion) { ' (exclusion)' })." -Level Detail

            # Hoisted out of the removal block below: the post-add verification
            # needs the same key, and computing it twice invites the two copies
            # drifting apart.
            $countKey = if ($Assignment.GroupID) { $Assignment.GroupID } else { $Assignment.Type }

            # Reset per assignment - these must not leak into the next iteration.
            $startDateTime = $null
            $deadlineDateTime = $null
            $useLocalTime = $false
            if ($Assignment.InstallTimeSettings) {
                $useLocalTime = [bool]$Assignment.InstallTimeSettings.useLocalTime
                $sourceStart = $Assignment.InstallTimeSettings.startDateTime
                $sourceDeadline = $Assignment.InstallTimeSettings.deadlineDateTime
                # Rebase onto the offset date, keeping the source time of day.
                # ConvertTo-YardstickWallClock reads the stored wall clock
                # without shifting it: Graph hands these back as Kind Utc under
                # PowerShell 7 and as a string under 5.1, and a bare [datetime]
                # cast of the string moves it into the host's timezone, so the
                # two disagree about what .Hour means.
                if ($null -ne $sourceStart) {
                    $sourceStart = ConvertTo-YardstickWallClock $sourceStart
                    $startDateTime = $AvailableDate.AddHours($sourceStart.Hour).AddMinutes($sourceStart.Minute)
                }
                if ($null -ne $sourceDeadline) {
                    $sourceDeadline = ConvertTo-YardstickWallClock $sourceDeadline
                    $deadlineDateTime = $DeadlineDate.AddHours($sourceDeadline.Hour).AddMinutes($sourceDeadline.Minute)
                }

                # Intune rejects a deadline that is already in the past unless
                # an available time accompanies it. Rebasing onto today (offset
                # 0) puts any morning deadline in the past for an afternoon run,
                # so nudge it forward. Intune treats a just-passed deadline the
                # same way: install at the next check-in.
                #
                # Compare against the clock the value is written in, not always
                # the local one: with useLocalTime false the wall clock is UTC,
                # and measuring it against local time misjudges "past" by the
                # host's offset in whichever direction the host sits.
                if (($null -ne $deadlineDateTime) -and ($null -eq $startDateTime)) {
                    $nowInDeadlineClock = if ($useLocalTime) { Get-Date } else { [datetime]::UtcNow }
                    if ($deadlineDateTime -lt $nowInDeadlineClock) {
                        $adjustedDeadline = [datetime]::SpecifyKind($nowInDeadlineClock.AddMinutes(5), [System.DateTimeKind]::Unspecified)
                        Write-Log -Level Warning "Rebased deadline $deadlineDateTime for $assignmentLabel is in the past; moving it to $adjustedDeadline so the assignment is still accepted."
                        $deadlineDateTime = $adjustedDeadline
                    }
                }

                Write-Log "Install time settings: UseLocalTime $useLocalTime, StartDateTime $startDateTime, DeadlineDateTime $deadlineDateTime." -Level Detail
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
                $addResult = $null
                $addWarnings = @()
                $addReturned = $false
                $conflicted = $false
                try {
                    # The Yardstick Intune cmdlets bail out of their Begin block
                    # with `break` (past deadline, expired token). A bare `break`
                    # from a called function is not catchable and unwinds to the
                    # caller's nearest enclosing loop - without this single-pass
                    # foreach to absorb it, one bad assignment would silently
                    # abandon every assignment after it. Absorbed here, it just
                    # leaves $addReturned false and is retried and logged.
                    foreach ($breakGuard in 1) {
                        if ($targetType -eq $allDevicesTarget) {
                            $addResult = Add-YardstickWin32AppAssignmentAllDevices @assignmentParams -WarningAction SilentlyContinue -WarningVariable addWarnings
                        }
                        elseif ($targetType -eq $allUsersTarget) {
                            $addResult = Add-YardstickWin32AppAssignmentAllUsers @assignmentParams -WarningAction SilentlyContinue -WarningVariable addWarnings
                        }
                        else {
                            $addResult = Add-YardstickWin32AppAssignmentGroup @assignmentParams -WarningAction SilentlyContinue -WarningVariable addWarnings
                        }
                        $addReturned = $true
                    }

                    # The cmdlets swallow every Graph failure into Write-Warning,
                    # so the warning stream is the only place the real error text
                    # exists. Left uncaptured it goes to the console and never
                    # reaches the log, which is how a BadRequest ended up
                    # recorded as a successful add.
                    foreach ($addWarning in $addWarnings) {
                        Write-Log "Intune warned while adding $assignmentLabel to $($To.id): $addWarning" -Level Warning
                        # Both the client-side duplicate guard and the server-side
                        # "The MobileApp Assignment already exists" BadRequest are
                        # idempotent - the assignment is already where we want it,
                        # and retrying can only reproduce the same conflict.
                        if ("$addWarning" -match 'already exists') { $conflicted = $true }
                    }

                    if (-not $addReturned) {
                        # A bail-out is not success, and must not reach the
                        # read-back: the assignment could be present on $To for
                        # unrelated reasons, and treating that as success here
                        # would delete the source copy.
                        Write-Log "Adding the $assignmentLabel assignment to $($To.id) aborted before returning. Check the preceding warning for the reason." -Level Warning
                    }
                    else {
                        # A returned assignment object is the one unambiguous
                        # success signal the cmdlets give us. Anything else -
                        # including a warning of any kind - has to be read back.
                        $successfullyAdded = if (($addWarnings.Count -eq 0) -and $addResult -and $addResult.id) {
                            $true
                        } else {
                            Test-YardstickAssignmentPresent -AppId $To.id -CountKey $countKey `
                                -Intent $Assignment.Intent -TargetType $Assignment.Type
                        }

                        if ($successfullyAdded) {
                            Write-Log "Added $assignmentLabel assignment to $($To.id)." -Level Detail
                        }
                        else {
                            Write-Log "$assignmentLabel assignment is not present on $($To.id) after the add." -Level Error
                        }
                    }
                }
                catch {
                    Write-Log "Could not add the $assignmentLabel assignment to $($To.id): $_" -Level Warning
                }

                if (!$successfullyAdded -and $conflicted) {
                    Write-Log "Not retrying $assignmentLabel on $($To.id): Intune reported a conflict that a retry cannot resolve." -Level Detail
                    break
                }
                if (!$successfullyAdded -and ($try -lt $maxRetries)) {
                    Write-Log "Retrying the $assignmentLabel assignment on $($To.id) (attempt $try/$maxRetries)..." -Level Detail
                    Start-Sleep -Seconds $RetryDelaySeconds
                }
            }

            # Remove the old assignment
            if ($successfullyAdded -and -not $CopyOnly) {
                if ([int]$targetCounts[$countKey] -gt 1) {
                    Write-Log -Level Warning "$($From.id) has $($targetCounts[$countKey]) assignments for $assignmentLabel. Removing one removes them all."
                }

                # No second verification pass: $successfullyAdded already means
                # Intune confirmed the assignment is on $To, either by returning
                # the created object or via the read-back above.
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
                                Remove-YardstickWin32AppAssignmentAllDevices -ID $From.id | Out-Null
                            }
                            elseif ($targetType -eq $allUsersTarget) {
                                Remove-YardstickWin32AppAssignmentAllUsers -ID $From.id | Out-Null
                            }
                            else {
                                Remove-YardstickWin32AppAssignmentGroup -ID $From.id -GroupID $Assignment.GroupID | Out-Null
                            }
                            $removeReturned = $true
                        }
                        if ($removeReturned) {
                            $successfullyRemoved = $true
                            Write-Log "Removed the $assignmentLabel assignment from $($From.id)." -Level Detail
                        }
                        else {
                            $removeFailure = "aborted before returning. Check the preceding warning for the reason"
                        }
                    }
                    catch {
                        $removeFailure = $_
                    }
                    if (-not $successfullyRemoved) {
                        Write-Log "Could not remove the $assignmentLabel assignment from $($From.id): $removeFailure" -Level Warning
                        if ($i -eq $maxRemovalAttempts) {
                            Write-Log "Gave up removing the $assignmentLabel assignment from $($From.id) after $maxRemovalAttempts attempts." -Level Error
                        }
                        else {
                            Write-Log "Retrying removal of the $assignmentLabel assignment from $($From.id) (attempt $($i + 1)/$maxRemovalAttempts)..." -Level Detail
                            Start-Sleep -Seconds $RetryDelaySeconds
                        }
                    }
                }
            }
            elseif (-not $successfullyAdded) {
                # The source assignment is the only surviving copy, so $From must
                # not be pruned out from under it.
                Write-Log -Level Error "$assignmentLabel assignment was not added to $($To.id); leaving the source assignment on $($From.id) intact and protecting it from deletion."
                & $addProtectedSourceId $From.id
            }
        }
    }
    if ($SkipDependencies) {
        Write-Log "Skipping dependency migration as requested." -Level Detail
        return
    }
    # Child dependencies are the apps that $From depends on; $To needs to depend
    # on the same apps. Add-YardstickWin32AppDependency REPLACES an app's entire
    # dependency set (it only preserves supersedence), so the complete desired
    # list has to be submitted in one call - adding them one at a time drops
    # every dependency configured by the previous call.
    if ($childDependencies.Count -gt 0) {
        # Apps the caller intends to delete this run. A dependency pointing at one
        # of them is a link that would block its deletion, so it is dropped rather
        # than carried forward - including one $To already holds, which Intune
        # would otherwise keep alive through the replace below.
        $excludedTargets = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($excludedId in $ExcludeDependencyTargetIds) {
            if ($excludedId) { [void]$excludedTargets.Add([string]$excludedId) }
        }

        # targetId -> dependencyType. Seed with what $To already depends on so
        # existing dependencies survive the replace.
        $desiredDependencies = [ordered]@{}
        foreach ($existing in (Get-YardstickWin32AppDependency -ID $To.id)) {
            if (($existing.PSObject.Properties.Name -contains "targetType") -and ($existing.targetType -eq "parent")) {
                continue
            }
            if ($excludedTargets.Contains([string]$existing.targetId)) {
                Write-Log "Dropping existing dependency $($existing.targetId) from $($To.id) because that app is scheduled for deletion." -Level Detail
                continue
            }
            $desiredDependencies[$existing.targetId] = ConvertTo-DependencyType $existing.dependencyType
        }
        foreach ($dependency in $childDependencies) {
            # An app cannot depend on itself, and $To must not inherit a
            # dependency on the app it is replacing.
            if (($dependency.targetId -eq $To.id) -or ($dependency.targetId -eq $From.id)) {
                Write-Log "Skipping child dependency $($dependency.targetId) because it points at the source or target app." -Level Detail
                continue
            }
            if ($excludedTargets.Contains([string]$dependency.targetId)) {
                Write-Log "Skipping child dependency $($dependency.targetId) because that app is scheduled for deletion." -Level Detail
                continue
            }
            $desiredDependencies[$dependency.targetId] = ConvertTo-DependencyType $dependency.dependencyType
        }

        if ($desiredDependencies.Count -eq 0) {
            Write-Log "No child dependencies to migrate to $($To.id)." -Level Detail
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
                        $dependencyObject = New-YardstickWin32AppDependency -ID $targetId -DependencyType $desiredDependencies[$targetId]
                        if ($dependencyObject) {
                            $dependencyObjects += $dependencyObject
                        }
                        else {
                            Write-Log "Could not build a dependency object for $targetId. It may no longer exist in Intune." -Level Warning
                        }
                    }
                    if ($dependencyObjects.Count -eq 0) {
                        Write-Log "None of the child dependencies could be resolved. Skipping dependency migration to $($To.id)." -Level Warning
                        break
                    }

                    Add-YardstickWin32AppDependency -ID $To.id -Dependency $dependencyObjects | Out-Null

                    # Add-YardstickWin32AppDependency warns instead of throwing when
                    # Graph rejects the update, so read the result back rather
                    # than assuming the call succeeded.
                    $assignedTargets = @(Get-YardstickWin32AppDependency -ID $To.id |
                        Where-Object { $_.targetType -ne "parent" } |
                        ForEach-Object { $_.targetId })
                    $missingTargets = @($dependencyObjects.targetId | Where-Object { $assignedTargets -notcontains $_ })
                    if ($missingTargets.Count -eq 0) {
                        $successfullyAdded = $true
                        Write-Log "Migrated $($dependencyObjects.Count) child dependency(ies) to $($To.id): $($dependencyObjects.targetId -join ', ')"
                    }
                    else {
                        Write-Log "Dependencies missing from $($To.id) after attempt $try/${maxRetries}: $($missingTargets -join ', ')" -Level Warning
                    }
                }
                catch {
                    Write-Log "Could not migrate child dependencies to $($To.id) (attempt $try/${maxRetries}): $_" -Level Warning
                }

                if (!$successfullyAdded -and ($try -lt $maxRetries)) {
                    Start-Sleep -Seconds 2
                }
            }
            if (!$successfullyAdded) {
                Write-Log "Could not migrate child dependencies to $($To.id) after $maxRetries attempts." -Level Error
            }
        }
    }

    if ($parentDependencies -and $parentDependencies.Count -gt 0) {
        $parentGroups = $parentDependencies | Group-Object -Property targetId
        if (-not $AllowDependentLinkUpdates) {
            Write-Log "Skipping dependent app link updates because AllowDependentLinkUpdates is disabled for this recipe." -Level Detail
            foreach ($group in $parentGroups) {
                $displayName = $group.Group[0].targetDisplayName
                if (-not $displayName) { $displayName = $group.Name }
                & $recordDependentStatus $displayName "Skipped (link updates disabled for recipe)"
            }
            & $addProtectedSourceId $From.id
        }
        elseif (-not [bool]$resolvedDependentOptions.Enabled) {
            Write-Log "Skipping dependent app link updates because dependentLinkUpdateEnabled is disabled in preferences." -Level Detail
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
                    $parentApp = Get-YardstickWin32App -Id $parentId
                    $parentDisplayName = $parentApp.DisplayName
                } catch {
                    Write-Log "Could not read the metadata of dependent app ${parentId}: $_" -Level Warning
                }
            }
            if (-not $parentDisplayName) {
                $parentDisplayName = $parentId
            }

            $normalizedParentName = $parentDisplayName.ToLowerInvariant()
            if ($normalizedBlocklist -and $normalizedBlocklist -contains $normalizedParentName) {
                Write-Log "Skipping dependent app $parentDisplayName because it is included in the blocklist."
                & $recordDependentStatus $parentDisplayName "Skipped (blocklisted)"
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
                    $parentDependencyList = Get-YardstickWin32AppDependency -ID $parentId
                    $childItems = $parentDependencyList | Where-Object { ($_.targetType -eq "child") -or (-not $_.targetType) }
                    if (-not $childItems) {
                        $statusMessage = "No dependencies to update"
                        $updateSucceeded = $true
                        break
                    }

                    $hasLinkToSource = $false
                    foreach ($entry in $childItems) {
                        $targetAppId = $entry.targetId
                        $normalizedTypeValue = ConvertTo-DependencyType $entry.dependencyType
                        $originalDependencies += New-YardstickWin32AppDependency -ID $targetAppId -DependencyType $normalizedTypeValue
                        if ($targetAppId -eq $From.id) {
                            $hasLinkToSource = $true
                            $targetAppId = $To.id
                        }
                        $updatedDependencies += New-YardstickWin32AppDependency -ID $targetAppId -DependencyType $normalizedTypeValue
                    }

                    if (-not $hasLinkToSource) {
                        $statusMessage = "Already up-to-date"
                        $updateSucceeded = $true
                        break
                    }

                    Remove-YardstickWin32AppDependency -ID $parentId | Out-Null
                    $dependenciesCleared = $true
                    Add-YardstickWin32AppDependency -ID $parentId -Dependency $updatedDependencies | Out-Null
                    $updateSucceeded = $true
                    $newTargetName = if ($To.DisplayName) { $To.DisplayName } else { $To.id }
                    Write-Log "Updated dependent app $parentDisplayName to reference $newTargetName."
                }
                catch {
                    $lastError = $_.Exception.Message
                    Write-Log "Failed to update dependent app $parentDisplayName on attempt $($attempt + 1): $lastError" -Level Warning
                    if ($dependenciesCleared -and $originalDependencies.Count -gt 0) {
                        try {
                            Add-YardstickWin32AppDependency -ID $parentId -Dependency $originalDependencies | Out-Null
                        } catch {
                            Write-Log "Could not restore the original dependencies of $parentDisplayName after the failure." -Level Error
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

    # The log report reads the display name from this line; keep its wording.
    Write-Log "Retrieving all versions of application with display name: $DisplayName" -Level Detail

    # Attempt to retrieve applications with retry logic
    $AllSimilarApps = Invoke-WithRetry -Label "Retrieve applications for $DisplayName" -MaxRetries 3 -DelaySeconds 5 -ScriptBlock {
        Get-YardstickWin32App -DisplayName "$DisplayName" -ErrorAction Stop
    } -OnFailure {
        Write-Log "Could not retrieve the versions of $DisplayName from Intune after 3 attempts. Exiting." -Level Error
        exit 1001
    }
    
    if (-not $AllSimilarApps) {
        Write-Log "No applications found with the name $DisplayName." -Level Detail
        return , @()
    }
    
    # Include the current name, any (N-x) rename, and the Ω DETECT - anchor for the same app.
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


function Get-NewestComparableVersion {
    <#
    .SYNOPSIS
    Returns the newest displayVersion from a set of existing apps that can actually
    be compared, or $null if there is none.

    .DESCRIPTION
    Callers must not reach for $ExistingVersions.displayVersion[0] directly. Two
    things go wrong with that:

    1. PowerShell unrolls a single-element property projection to a scalar, so when
       exactly one matching app exists, [0] indexes into the version *string* and
       returns its first character ("2" instead of "2.10.91.91").
    2. An app can have a null or empty displayVersion -- for example one renamed to
       an "Ω DETECT -" anchor, or created by hand -- and feeding that to
       Compare-AppVersions throws, which fails the entire recipe.

    This takes the already newest-first list from Get-SameAppAllVersions and returns
    the first entry that carries a usable version.

    .PARAMETER ExistingVersions
    Applications sorted newest first, as returned by Get-SameAppAllVersions.

    .OUTPUTS
    String version, or $null when no entry has a usable displayVersion.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyCollection()]
        $ExistingVersions
    )

    if ($null -eq $ExistingVersions) { return $null }

    $newest = @($ExistingVersions) |
        Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace($_.displayVersion) } |
        Select-Object -First 1

    if ($null -eq $newest) { return $null }
    return [string]$newest.displayVersion
}






function Test-IsVersionDetection {
    <#
    .SYNOPSIS
    Returns $true when the detection type is one that compares against a version
    number and therefore benefits from a Ω DETECT - anchor to catch older installs.
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


function New-YardstickDetectionRuleFromSpec {
    <#
    .SYNOPSIS
    Builds one Intune file, msi, or registry detection rule from recipe fields.

    .DESCRIPTION
    The spec is keyed by the recipe field names (detectionType, fileDetection*,
    registryDetection*), so a recipe's primary rule and each entry in its
    additionalDetectionRules list are built by the same code. Script detection
    is not handled here: it writes a file to disk, and Intune does not allow a
    script rule to be combined with any other rule.

    .PARAMETER Spec
    Dictionary of recipe detection fields describing a single rule.

    .PARAMETER ProductCode
    The MSI product code (required for msi rules).

    .PARAMETER DefaultVersion
    Used when a version comparison (file or registry) has no explicit value.

    .PARAMETER Is32BitApp
    Sets check32BitOn64System on registry comparison rules.
    #>
    param(
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Spec,
        [string]$ProductCode = "",
        [string]$DefaultVersion = "",
        [bool]$Is32BitApp = $false
    )

    $fileVersion = if ($Spec['fileDetectionVersion']) { [string]$Spec['fileDetectionVersion'] } else { $DefaultVersion }
    $path = [string]$Spec['fileDetectionPath']
    $name = [string]$Spec['fileDetectionName']
    $fileOperator = [string]$Spec['fileDetectionOperator']
    $keyPath = [string]$Spec['registryDetectionKey']
    $valueName = [string]$Spec['registryDetectionValueName']
    $registryValue = $Spec['registryDetectionValue']
    $registryOperator = [string]$Spec['registryDetectionOperator']

    switch ([string]$Spec['detectionType']) {
        "file" {
            switch ([string]$Spec['fileDetectionMethod']) {
                "exists" {
                    return New-YardstickWin32AppDetectionRuleFile -Existence -DetectionType "exists" -Path $path -FileOrFolder $name
                }
                "modified" {
                    return New-YardstickWin32AppDetectionRuleFile -DateModified -Path $path -FileOrFolder $name -Operator $fileOperator -DateTimeValue $Spec['fileDetectionDateTime']
                }
                "created" {
                    return New-YardstickWin32AppDetectionRuleFile -DateCreated -Path $path -FileOrFolder $name -Operator $fileOperator -DateTimeValue (Get-Date $Spec['fileDetectionDateTime'])
                }
                "version" {
                    return New-YardstickWin32AppDetectionRuleFile -Version -Path $path -FileOrFolder $name -Operator $fileOperator -VersionValue $fileVersion
                }
                "size" {
                    return New-YardstickWin32AppDetectionRuleFile -Size -Path $path -FileOrFolder $name -Operator $fileOperator -SizeinMBValue $Spec['fileDetectionValue']
                }
                default {
                    throw "Unsupported fileDetectionMethod '$($Spec['fileDetectionMethod'])'"
                }
            }
        }
        "msi" {
            return New-YardstickWin32AppDetectionRuleMsi -ProductCode $ProductCode -ProductVersion $fileVersion
        }
        "registry" {
            switch ([string]$Spec['registryDetectionMethod']) {
                "exists" {
                    if ($registryValue) {
                        return New-YardstickWin32AppDetectionRuleRegistry -Existence -KeyPath $keyPath -ValueName $valueName -DetectionType "exists"
                    } else {
                        return New-YardstickWin32AppDetectionRuleRegistry -Existence -KeyPath $keyPath -DetectionType "exists"
                    }
                }
                "version" {
                    $versionValue = if ($registryValue) { [string]$registryValue } else { $DefaultVersion }
                    return New-YardstickWin32AppDetectionRuleRegistry -VersionComparison -KeyPath $keyPath -ValueName $valueName -Check32BitOn64System $Is32BitApp -VersionComparisonOperator $registryOperator -VersionComparisonValue $versionValue
                }
                "integer" {
                    return New-YardstickWin32AppDetectionRuleRegistry -IntegerComparison -KeyPath $keyPath -ValueName $valueName -Check32BitOn64System $Is32BitApp -IntegerComparisonOperator $registryOperator -IntegerComparisonValue $registryValue
                }
                "string" {
                    return New-YardstickWin32AppDetectionRuleRegistry -StringComparison -KeyPath $keyPath -ValueName $valueName -Check32BitOn64System $Is32BitApp -StringComparisonOperator $registryOperator -StringComparisonValue $registryValue
                }
                default {
                    throw "Unsupported registryDetectionMethod '$($Spec['registryDetectionMethod'])'"
                }
            }
        }
        default {
            throw "Unsupported detectionType '$($Spec['detectionType'])' - only file, msi, and registry rules can be built from a spec"
        }
    }
}


function Get-DetectAnchor {
    <#
    .SYNOPSIS
    Returns the "Ω DETECT - <DisplayName>" Intune app object if one exists, else $null.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$DisplayName
    )
    $anchorName = Get-DetectAnchorName -DisplayName $DisplayName
    $candidates = Get-YardstickWin32App -DisplayName $anchorName -ErrorAction SilentlyContinue
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
    return "Ω DETECT - $DisplayName"
}


function Set-DetectAnchor {
    <#
    .SYNOPSIS
    Renames the given Intune app to "Ω DETECT - <DisplayName>" so it is preserved
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
        Write-Log -Level Detail "App $($App.Id) is already the Ω DETECT - anchor for $DisplayName"
        return
    }
    Write-Log "Pinning $($App.DisplayName) ($($App.Id)) as Ω DETECT - anchor for $DisplayName"
    Set-YardstickWin32App -Id $App.Id -DisplayName $anchorName | Out-Null
}


function Get-YardstickSupersedenceRelationship {
    <#
    .SYNOPSIS
    Returns the supersedence relationships that touch the given Win32 app, in both
    directions.

    .DESCRIPTION
    The beta `mobileApps/{id}/relationships` collection contains an entry for every
    supersedence edge the app participates in, reported from the perspective of the
    app being queried: `sourceId` is always the queried app and `targetId` the app
    at the other end, in BOTH directions. `targetType` is what names the direction -
    'child' means this app supersedes the target, 'parent' means the target
    supersedes this app. Reverse links must be cleared from the *parent* before
    Intune will allow the child to be deleted.

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
        Write-Log "Could not read the supersedence relationships of ${Id}: $_" -Level Warning
        return @()
    }

    # Returns are NOT comma-wrapped. `return ,@($x)` emits the outer one-element
    # array, which unrolls to hand the caller a single item that is itself the
    # array - so `@(Get-YardstickSupersedenceRelationship ...).Count` came back as
    # 1 for every real count, including 0 and 2. That is what produced
    # "Expected 2 supersedence target(s) ... but Intune reports 1" even on runs
    # where Intune had stored both, and it made the stale-link strip in
    # Set-YardstickSupersedence fire against apps with no links at all.
    #
    # The unrolling this guarded against is real - a bare [PSCustomObject] has no
    # synthetic .Count under Windows PowerShell 5.1 - but every caller already
    # wraps the call in @(), which normalizes 0, 1 and many correctly. Keep it
    # that way: call this as @(Get-YardstickSupersedenceRelationship ...).
    #
    # Direction comes off targetType, NOT sourceId. Graph reports this collection
    # from the perspective of the app being queried, so sourceId equals $Id for
    # forward and reverse links alike - querying a superseded child returns
    # sourceId=<child>, targetId=<parent>, targetType='parent'. Keying off sourceId
    # therefore classified every link as Forward and left Reverse permanently
    # empty, so Clear-YardstickAppLink never detached a superseded app from its
    # parent; it cleared the child's own (empty) forward set instead, the read-back
    # still saw the link, and the prune failed with "supersedes <parent id>".
    # That is what left stale (N-2) versions behind.
    #
    # Payloads built locally and not yet read back from Graph carry no targetType,
    # so fall back to the sourceId comparison for those.
    $isForward = {
        param($relationship)
        if ($relationship.targetType) { return ($relationship.targetType -eq 'child') }
        return ((-not $relationship.sourceId) -or ($relationship.sourceId -eq $Id))
    }
    switch ($Direction) {
        'Forward' { return @($relationships | Where-Object { & $isForward $_ }) }
        'Reverse' { return @($relationships | Where-Object { -not (& $isForward $_) }) }
        default   { return $relationships }
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
        Write-Log "Clearing supersedence on $ParentId (it was the only target: $TargetId)." -Level Detail
        Remove-YardstickWin32AppSupersedence -ID $ParentId | Out-Null
        return
    }

    $rebuilt = New-SupersedenceObject -Relationships $remaining
    if ($rebuilt.Count -eq 0) {
        Write-Log "Could not rebuild supersedence for $ParentId; leaving it untouched." -Level Warning
        return
    }
    Write-Log "Rebuilding supersedence on $ParentId without target $TargetId ($($rebuilt.Count) remaining)." -Level Detail
    Add-YardstickWin32AppSupersedence -ID $ParentId -Supersedence $rebuilt | Out-Null
}


function ConvertTo-DependencyType {
    <#
    .SYNOPSIS
    Maps a dependencyType off a Graph relationship onto the casing
    New-YardstickWin32AppDependency validates.

    .DESCRIPTION
    Graph hands the type back lowercase ("autoinstall"), and the cmdlet's
    ValidateSet only accepts "AutoInstall"/"Detect". Anything unrecognised - or
    missing entirely - falls back to Detect, which is the weaker of the two: it
    reports the dependency as unmet rather than silently installing an app the
    admin never asked for.
    #>
    param(
        [AllowNull()]
        $DependencyType
    )

    if (-not $DependencyType) { return "Detect" }
    switch ($DependencyType.ToString().ToLower()) {
        "autoinstall" { return "AutoInstall" }
        default       { return "Detect" }
    }
}


function Remove-DependencyReference {
    <#
    .SYNOPSIS
    Removes a single dependency target from a parent app, leaving the parent's
    other dependencies (and its supersedence) intact.

    .DESCRIPTION
    The dependency mirror of Remove-SupersedenceReference. Intune has no "delete
    one relationship" operation for Win32 apps - Add-YardstickWin32AppDependency
    replaces the whole dependency set (preserving supersedence) - so this rebuilds
    the parent's child dependency list without $TargetId and re-submits it.

    Needed because Intune refuses to delete an app that anything still depends on,
    and Remove-YardstickApp previously only unwound supersedence.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ParentId,

        [Parameter(Mandatory = $true)]
        [string]$TargetId
    )

    # A parent's own dependency list is its child entries; entries flagged
    # "parent" describe apps that depend on it and are not ours to rewrite.
    $children = @(Get-YardstickWin32AppDependency -ID $ParentId |
        Where-Object { ($_.targetType -eq "child") -or (-not $_.targetType) })
    if (-not ($children | Where-Object targetId -eq $TargetId)) {
        return
    }

    $remaining = @($children | Where-Object targetId -ne $TargetId)
    if ($remaining.Count -eq 0) {
        Write-Log "Clearing dependencies on $ParentId (it was the only target: $TargetId)." -Level Detail
        Remove-YardstickWin32AppDependency -ID $ParentId | Out-Null
        return
    }

    $rebuilt = @()
    foreach ($entry in $remaining) {
        $dependencyObject = New-YardstickWin32AppDependency -ID $entry.targetId `
            -DependencyType (ConvertTo-DependencyType $entry.dependencyType)
        if ($dependencyObject) {
            $rebuilt += $dependencyObject
        }
        else {
            Write-Log "Could not rebuild the dependency on $ParentId for target $($entry.targetId). It may no longer exist in Intune." -Level Warning
        }
    }

    if ($rebuilt.Count -eq 0) {
        # Every survivor failed to resolve. Clearing outright is still correct -
        # the caller needs $TargetId detached and the leftovers point at apps
        # Intune can no longer find.
        Write-Log "Clearing dependencies on ${ParentId}: none of the $($remaining.Count) remaining target(s) could be rebuilt." -Level Warning
        Remove-YardstickWin32AppDependency -ID $ParentId | Out-Null
        return
    }

    Write-Log "Rebuilding dependencies on $ParentId without target $TargetId ($($rebuilt.Count) remaining)." -Level Detail
    Add-YardstickWin32AppDependency -ID $ParentId -Dependency $rebuilt | Out-Null
}


function New-SupersedenceObject {
    <#
    .SYNOPSIS
    Builds the OrderedDictionary array that Add-YardstickWin32AppSupersedence expects.

    .DESCRIPTION
    Accepts either existing relationship objects (which carry targetId +
    supersedenceType) or an explicit target id / type pair. Any entry that cannot
    be resolved is dropped rather than passed through as $null, because
    Add-YardstickWin32AppSupersedence declares [OrderedDictionary[]] with
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
            # supersedenceType comes back lowercase from Graph; New-YardstickWin32AppSupersedence validates Update/Replace.
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
        $object = New-YardstickWin32AppSupersedence -ID $pair.TargetId -SupersedenceType $pair.Type
        if ($object) {
            $built.Add([System.Collections.Specialized.OrderedDictionary]$object) | Out-Null
        } else {
            Write-Log "Could not build a supersedence object for target $($pair.TargetId). Skipping it." -Level Warning
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
    Ω DETECT - anchor lives here: its detection rule deliberately matches a very wide
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
        Write-Log "No supersedence targets for $($NewApp.DisplayName). Clearing any stale links."
        if (@(Get-YardstickSupersedenceRelationship -Id $NewApp.id -Direction Forward).Count -gt 0) {
            Remove-YardstickWin32AppSupersedence -ID $NewApp.id | Out-Null
        }
        return 0
    }

    $updateOnly = [System.Collections.Generic.HashSet[string]]::new([string[]]$UpdateOnlyIds, [StringComparer]::OrdinalIgnoreCase)

    # Intune caps a supersedence graph at 10 nodes, one of which is the parent.
    # Update-only targets (the Ω DETECT - anchor) are reserved first: the anchor is
    # by definition the oldest version, so a plain newest-first trim would drop
    # exactly the target that catches stale installs.
    $maxTargets = 9
    if ($targets.Count -gt $maxTargets) {
        Write-Log "$($targets.Count) supersedence targets exceed Intune's limit of ${maxTargets}. Trimming the oldest." -Level Warning
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
                Write-Log "Clearing stale supersedence on $($target.DisplayName) ($($target.id))." -Level Detail
                Remove-YardstickWin32AppSupersedence -ID $target.id | Out-Null
            }
        } catch {
            Write-Log "Could not clear the existing supersedence on $($target.DisplayName) ($($target.id)): $_" -Level Warning
        }
    }

    $supersedence = [System.Collections.Generic.List[System.Collections.Specialized.OrderedDictionary]]::new()
    foreach ($target in $targets) {
        $targetType = if ($updateOnly.Contains($target.id)) { 'Update' } else { $Type }
        $built = New-SupersedenceObject -TargetIds @($target.id) -Type $targetType
        foreach ($object in $built) { $supersedence.Add($object) | Out-Null }
    }

    if ($supersedence.Count -eq 0) {
        Write-Log "No resolvable supersedence targets for $($NewApp.DisplayName). Skipping supersedence." -Level Warning
        return 0
    }

    # The log report matches this line; keep its wording (and no trailing period).
    Write-Log "Attaching supersedence ($Type) from $($NewApp.DisplayName) to $($supersedence.Count) target(s)"
    # Add-YardstickWin32AppSupersedence replaces the parent's whole supersedence set,
    # so there is no need to clear it first.
    Add-YardstickWin32AppSupersedence -ID $NewApp.id -Supersedence $supersedence.ToArray() | Out-Null

    # The cmdlet downgrades Graph failures to warnings, so read the graph back and
    # report what Intune actually stored rather than what we asked for.
    $attached = @(Get-YardstickSupersedenceRelationship -Id $NewApp.id -Direction Forward)
    if ($attached.Count -ne $supersedence.Count) {
        Write-Log "Expected $($supersedence.Count) supersedence target(s) on $($NewApp.DisplayName), but Intune reports $($attached.Count)." -Level Error
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
    Get-YardstickWin32AppAssignment because that cmdlet does not surface the assignment
    id, which is required to PATCH an individual assignment.

    .PARAMETER SkipGroupIds
    Entra group ids that must never auto-update. Assignments targeting one of these
    groups are forced to 'notConfigured' rather than merely left alone, so adding a
    group to the list turns off auto-update a previous run had already enabled.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$AppId,
        [bool]$Enabled = $true,
        [ValidateSet('', 'required', 'available')]
        [string]$IntentFilter = 'available',
        [string[]]$SkipGroupIds = @()
    )

    # Group ids come from YAML, where casing is whatever the admin pasted in.
    $skipSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($groupId in $SkipGroupIds) {
        if ($groupId) { $skipSet.Add([string]$groupId) | Out-Null }
    }

    try {
        $assignments = @(Invoke-YardstickGraphRequest -Resource "deviceAppManagement/mobileApps/$AppId/assignments")
    } catch {
        Write-Log "Could not read the assignments of $AppId while setting auto-update: $_" -Level Warning
        return 0
    }

    if ($IntentFilter) {
        $assignments = @($assignments | Where-Object intent -eq $IntentFilter)
    }
    if ($assignments.Count -eq 0) {
        Write-Log "No $(if ($IntentFilter) { "$IntentFilter-intent " })assignments on $AppId to configure auto-update on."
        return 0
    }

    $count = 0
    $skipped = 0
    foreach ($assignment in $assignments) {
        if (-not $assignment.id) { continue }

        # Exclusion targets carry no settings object and cannot auto-update.
        if ($assignment.target.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget') { continue }

        # A skipped group is forced off regardless of $Enabled. All-devices and
        # all-users targets carry no groupId and can never be skipped this way.
        $isSkipped = $assignment.target.groupId -and $skipSet.Contains([string]$assignment.target.groupId)
        $desiredState = if ($Enabled -and -not $isSkipped) { 'enabled' } else { 'notConfigured' }

        if ($assignment.settings.autoUpdateSettings.autoUpdateSupersededAppsState -eq $desiredState) {
            if ($isSkipped) { $skipped++ } else { $count++ }
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
            if ($isSkipped) { $skipped++ } else { $count++ }
        } catch {
            Write-Log "Could not set auto-update on assignment $($assignment.id) of ${AppId}: $_" -Level Warning
        }
    }
    $desiredStateForApp = if ($Enabled) { 'enabled' } else { 'notConfigured' }
    Write-Log "Auto-update ($desiredStateForApp) is set on $count assignment(s) for $AppId."
    if ($skipped -gt 0) {
        Write-Log "Auto-update held at notConfigured on $skipped assignment(s) for $AppId (group in the auto-update skip list)."
    }
    return $count
}


function Clear-YardstickAppLink {
    <#
    .SYNOPSIS
    Removes every link that would stop Intune deleting a Win32 app, and reports
    anything it could not remove.

    .DESCRIPTION
    Intune refuses to delete an app that still participates in a relationship, and
    Remove-YardstickWin32App downgrades that refusal to a warning - so a prune used to
    fail with no indication of which link was responsible. This strips all of them,
    in the order that leaves the app least exposed if a later step fails:

      1. Assignments                - devices stop being targeted before the
                                      relationship graph is torn down
      2. Dependencies, reverse      - apps that depend on this one
      3. Dependencies, forward      - apps this one depends on
      4. Supersedence, reverse      - apps that supersede this one
      5. Supersedence, forward      - apps this one supersedes

    Individual assignments are deleted through Graph by assignment id. The
    Yardstick Intune module can only remove assignments by target, which takes out
    every assignment sharing that target - see the note in
    Move-AssignmentsAndDependencies. They are all going anyway, but by-id keeps the
    log honest about what was actually removed.

    Assignments are read straight from Graph rather than via
    Get-YardstickWin32AppAssignment, which returns $null for an app holding exactly one
    assignment under Windows PowerShell 5.1 (see the note in
    Move-AssignmentsAndDependencies) and does not surface the assignment id anyway.

    Call this immediately before the delete, never earlier: an app whose links were
    stripped and which then survives is worse off than one left alone. No single
    stuck link abandons the rest, and the closing sweep re-reads Intune so the
    return value reflects what is really still there - a read that could not be
    completed counts as a surviving link rather than being taken for "clear".

    .PARAMETER RetryDelaySeconds
    Seconds between attempts (default: 2). Exists so tests can drive the retry paths
    without sleeping.

    .OUTPUTS
    String array describing the links that survived. Empty means the app is
    deletable. Call as @(Clear-YardstickAppLink ...) - an empty array unrolls to
    $null otherwise.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $true)]
        $App,

        [int]$RetryDelaySeconds = 2
    )

    $appId = $App.id
    $label = if ($App.DisplayName) { "$($App.DisplayName) ($appId)" } else { $appId }

    # Every YardstickWin32App cmdlet can bail out of its Begin block with a bare
    # `break` (expired token, most often) rather than throwing. A bare break is not
    # catchable and unwinds to the caller's nearest enclosing loop - see the long
    # note in Move-AssignmentsAndDependencies. One escaping from here would skip
    # both the unwind and the verification below, and Remove-YardstickApp would go
    # on to delete an app whose links were never touched. So every read runs inside
    # a single-pass foreach that absorbs it and reports Ok = $false, which the
    # closing sweep treats as "not proven deletable" rather than "clear".
    $readLinks = {
        param([string]$What, [scriptblock]$Read)
        $items = @()
        $ok = $false
        try {
            foreach ($breakGuard in 1) {
                $items = @(& $Read)
                $ok = $true
            }
            if (-not $ok) {
                Write-Log "Reading $What on $label aborted before returning. Check the preceding warning for the reason." -Level Warning
            }
        } catch {
            Write-Log "Could not read $What on ${label}: $_" -Level Warning
        }
        [PSCustomObject]@{ Ok = $ok; Items = $items; What = $What }
    }

    # Each read is used twice - once to decide what to remove, once to verify the
    # removal took - so the queries live in one place.
    $assignmentQuery = {
        Invoke-YardstickGraphRequest -Resource "deviceAppManagement/mobileApps/$appId/assignments"
    }
    $dependentQuery = {
        Get-YardstickWin32AppDependency -ID $appId |
            Where-Object targetType -eq "parent" |
            Select-Object -ExpandProperty targetId -Unique
    }
    $dependencyQuery = {
        Get-YardstickWin32AppDependency -ID $appId |
            Where-Object { ($_.targetType -eq "child") -or (-not $_.targetType) } |
            Select-Object -ExpandProperty targetId -Unique
    }
    $supersedenceQuery = {
        Get-YardstickSupersedenceRelationship -Id $appId
    }

    # 1. Assignments.
    foreach ($assignment in (& $readLinks "assignments" $assignmentQuery).Items) {
        if (-not $assignment.id) { continue }
        $assignmentId = $assignment.id
        $target = if ($assignment.target.groupId) {
            "group $($assignment.target.groupId)"
        } else {
            $assignment.target.'@odata.type'
        }
        $removed = Invoke-WithRetry -Label "Remove assignment for $target from $label" `
            -MaxRetries 3 -DelaySeconds $RetryDelaySeconds -ScriptBlock {
                # Graph throws on failure here, unlike the module's Remove-*
                # cmdlets, so a return is proof enough - no VerifyBlock needed.
                Invoke-YardstickGraphRequest -Method Delete `
                    -Resource "deviceAppManagement/mobileApps/$appId/assignments/$assignmentId" | Out-Null
                $true
            }
        if ($removed) {
            Write-Log "Removed $($assignment.intent) assignment for $target from $label." -Level Detail
        }
    }

    # 2. Apps that depend on this one. Intune will not delete a dependency target.
    #    Invoke-WithRetry's own retry loop absorbs a break out of the cmdlets here
    #    and reports the attempt as failed, which is what we want.
    foreach ($dependentId in (& $readLinks "dependent apps" $dependentQuery).Items) {
        $currentDependentId = $dependentId
        $detached = Invoke-WithRetry -Label "Detach $label from dependent app $currentDependentId" `
            -MaxRetries 3 -DelaySeconds $RetryDelaySeconds -ScriptBlock {
                Remove-DependencyReference -ParentId $currentDependentId -TargetId $appId
                $true
            } -VerifyBlock {
                -not @(Get-YardstickWin32AppDependency -ID $currentDependentId |
                    Where-Object { (($_.targetType -eq "child") -or (-not $_.targetType)) -and ($_.targetId -eq $appId) })
            }
        if ($detached) {
            Write-Log "Detached $label from dependent app $currentDependentId." -Level Detail
        }
    }

    # 3. This app's own dependencies.
    $dependencyIds = @((& $readLinks "dependencies" $dependencyQuery).Items)
    if ($dependencyIds.Count -gt 0) {
        $cleared = Invoke-WithRetry -Label "Clear $($dependencyIds.Count) dependency(ies) from $label" `
            -MaxRetries 3 -DelaySeconds $RetryDelaySeconds -ScriptBlock {
                Remove-YardstickWin32AppDependency -ID $appId | Out-Null
                $true
            } -VerifyBlock {
                $recheck = & $readLinks "dependencies" $dependencyQuery
                $recheck.Ok -and ($recheck.Items.Count -eq 0)
            }
        if ($cleared) {
            Write-Log "Cleared dependencies from $label ($($dependencyIds -join ', '))." -Level Detail
        }
    }

    # 4/5. Supersedence, both directions. Graph reports this collection from the
    #      queried app's perspective - sourceId is $appId on forward AND reverse
    #      links alike, and targetType names the direction - so the app at the
    #      other end is whichever id is not $appId.
    $isReverseLink = {
        param($relationship)
        if ($relationship.targetType) { return ($relationship.targetType -eq 'parent') }
        return ($relationship.sourceId -and ($relationship.sourceId -ne $appId))
    }
    $otherEnd = {
        param($relationship)
        if ($relationship.sourceId -and ($relationship.sourceId -ne $appId)) {
            $relationship.sourceId
        } else {
            $relationship.targetId
        }
    }

    $relationships = @((& $readLinks "supersedence" $supersedenceQuery).Items)

    $supersedingIds = @($relationships |
        Where-Object { & $isReverseLink $_ } |
        ForEach-Object { & $otherEnd $_ } |
        Where-Object { $_ } |
        Select-Object -Unique)
    foreach ($supersedingId in $supersedingIds) {
        $currentSupersedingId = $supersedingId
        Invoke-WithRetry -Label "Detach $label from superseding app $currentSupersedingId" `
            -MaxRetries 3 -DelaySeconds $RetryDelaySeconds -ScriptBlock {
                Remove-SupersedenceReference -ParentId $currentSupersedingId -TargetId $appId
                $true
            } | Out-Null
    }

    $hasForwardLinks = @($relationships | Where-Object { -not (& $isReverseLink $_) }).Count -gt 0
    if ($hasForwardLinks) {
        Invoke-WithRetry -Label "Strip supersedence from $label" `
            -MaxRetries 3 -DelaySeconds $RetryDelaySeconds -ScriptBlock {
                Remove-YardstickWin32AppSupersedence -ID $appId | Out-Null
                $true
            } | Out-Null
    }

    # 6. Read everything back. The cmdlets above downgrade Graph failures to
    #    warnings, so only a re-read can say whether the app is really deletable -
    #    and a read that could not be completed counts against it.
    $surviving = [System.Collections.Generic.List[string]]::new()

    foreach ($query in @(
        @{ What = "assignments";  Query = $assignmentQuery;   Describe = { param($x) "assignment $($x.id)" } }
        @{ What = "dependent apps"; Query = $dependentQuery;  Describe = { param($x) "dependency from $x" } }
        @{ What = "dependencies"; Query = $dependencyQuery;   Describe = { param($x) "dependency on $x" } }
    )) {
        $result = & $readLinks $query.What $query.Query
        if (-not $result.Ok) {
            $surviving.Add("$($query.What) (could not be read back)")
            continue
        }
        foreach ($item in $result.Items) {
            $surviving.Add((& $query.Describe $item))
        }
    }

    $supersedenceResult = & $readLinks "supersedence" $supersedenceQuery
    if (-not $supersedenceResult.Ok) {
        $surviving.Add("supersedence (could not be read back)")
    }
    foreach ($relationship in $supersedenceResult.Items) {
        if (& $isReverseLink $relationship) {
            $surviving.Add("superseded by $(& $otherEnd $relationship)")
        } else {
            $surviving.Add("supersedes $($relationship.targetId)")
        }
    }

    return [string[]]$surviving
}


function Remove-YardstickApp {
    <#
    .SYNOPSIS
    Safely deletes an Intune Win32 app.

    .DESCRIPTION
    Intune refuses to delete an app that still participates in a relationship, so
    Clear-YardstickAppLink unwinds all of them - assignments, dependencies and
    supersedence, in both directions - immediately before the delete. When a link
    cannot be removed the delete is not attempted at all: it could only fail, and
    throwing here names the link instead of leaving the caller with a bare
    "failed to remove". Callers treat that throw as a failed prune, which in
    Yardstick.ps1 means the app is superseded rather than deleted.
    #>
    param(
        [Parameter(Mandatory=$true)]
        $App
    )

    $surviving = @(Clear-YardstickAppLink -App $App)
    if ($surviving.Count -gt 0) {
        throw "Cannot delete $($App.DisplayName) ($($App.id)): $($surviving.Count) link(s) could not be removed - $($surviving -join '; ')"
    }

    Write-Log "Removing $($App.DisplayName) ($($App.id))..."

    # Intune is eventually consistent about relationship teardown: Clear-YardstickAppLink
    # can report every link removed and the delete still be refused a second later with
    # "Cannot delete this app at this time. Try again shortly." Invoke-YardstickGraphRequest
    # only retries 408/429/5xx, so this 4xx would otherwise fail the prune outright.
    $deleteBackoff = @(5, 10, 15)
    for ($attempt = 0; ; $attempt++) {
        try {
            Remove-YardstickWin32App -Id $App.id
            break
        } catch {
            $message = [string]$_
            $isRelationshipLag = $message -match 'Cannot delete this app at this time' -or
                                 $message -match 'MetadataBehaviorValidateRelationshipsForAppToDelete'
            if (-not $isRelationshipLag -or $attempt -ge $deleteBackoff.Count) { throw }
            $delay = $deleteBackoff[$attempt]
            Write-Log "Intune is not ready to delete $($App.DisplayName) ($($App.id)) yet. Retrying in $delay second(s) (attempt $($attempt + 1)/$($deleteBackoff.Count + 1))..." -Level Detail
            Start-Sleep -Seconds $delay
        }
    }

    # Remove-YardstickWin32App downgrades Graph failures (including Intune's refusal
    # to delete an app that is still in a relationship) to a warning, so confirm
    # the app is really gone. Callers rely on this throwing to know a prune failed.
    $stillPresent = $null
    try {
        $stillPresent = Get-YardstickWin32App -Id $App.id -ErrorAction SilentlyContinue
    } catch {
        $stillPresent = $null
    }
    if ($stillPresent) {
        throw "Intune still reports app $($App.DisplayName) ($($App.id)) after the delete request."
    }
}



function Get-YardstickDropboxPayload {
    <#
    .SYNOPSIS
    Returns the installer files an operator has staged for a manual-drop recipe.

    .DESCRIPTION
    Some vendors put their installer behind a signed-in, licensed download that
    Yardstick cannot automate (CLO, and anything else gated on an entitlement
    rather than a session cookie). Those recipes set `manualDownload: true`, and
    an operator drops the payload into <DropboxRoot>\<Folder>.

    An absent or empty folder is the normal steady state - Complete-YardstickDropbox
    empties it after every successful upload - so it returns $null rather than
    throwing. Only a missing dropbox configuration is an error.

    .PARAMETER DropboxRoot
    The SoftwareDropbox root from preferences.yaml.

    .PARAMETER Folder
    Per-recipe subfolder, normally the recipe id.

    .OUTPUTS
    PSCustomObject with Path and Files, or $null when nothing is staged.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][AllowNull()][string]$DropboxRoot,
        [Parameter(Mandatory)][string]$Folder
    )

    if ([string]::IsNullOrWhiteSpace($DropboxRoot)) {
        throw "Recipe sets manualDownload but no SoftwareDropbox path is configured in preferences.yaml."
    }

    $path = Join-Path $DropboxRoot $Folder
    if (-not (Test-Path -LiteralPath $path)) { return $null }

    $files = @(Get-ChildItem -LiteralPath $path -Force | Where-Object Name -ne '.gitkeep')
    if ($files.Count -eq 0) { return $null }

    return [PSCustomObject]@{
        Path  = $path
        Files = $files
    }
}


function Complete-YardstickDropbox {
    <#
    .SYNOPSIS
    Archives a manual-drop payload after a successful upload, emptying the dropbox.

    .DESCRIPTION
    Moves everything in the dropbox folder to <ArchiveRoot>\<Folder>\<Version>.
    Emptying the dropbox is what makes the next run a no-op: a manual-drop recipe
    only packages when an operator has staged something new. This runs only after
    Intune has accepted the upload, so a failed run leaves the payload in place to
    be retried.

    .PARAMETER DropboxPath
    The folder returned by Get-YardstickDropboxPayload.

    .PARAMETER ArchiveRoot
    The SoftwareArchive root from preferences.yaml.

    .PARAMETER Folder
    Per-recipe subfolder, normally the recipe id.

    .PARAMETER Version
    The version that was published, used as the archive subfolder name.
    #>
    param(
        [Parameter(Mandatory)][string]$DropboxPath,
        [Parameter(Mandatory)][AllowEmptyString()][AllowNull()][string]$ArchiveRoot,
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$Version
    )

    if ([string]::IsNullOrWhiteSpace($ArchiveRoot)) {
        Write-Log "No SoftwareArchive path is configured. Leaving the payload in $DropboxPath." -Level Warning
        return
    }

    $destination = Join-Path (Join-Path $ArchiveRoot $Folder) $Version
    if (Test-Path -LiteralPath $destination) {
        # Same version staged twice - a -Force run, or a corrected payload. Keep
        # both rather than failing the move.
        $destination = "$destination-$(Get-Date -Format 'yyyyMMddHHmmss')"
    }
    New-Item -ItemType Directory -Path $destination -Force | Out-Null

    foreach ($item in @(Get-ChildItem -LiteralPath $DropboxPath -Force)) {
        Move-Item -LiteralPath $item.FullName -Destination $destination -Force
    }
    Write-Log "Archived the manual drop to $destination."
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


function Import-RecipeOverrides {
    <#
    .SYNOPSIS
    Reads the site's recipe overrides file (Local\recipeOverrides.yaml).

    .DESCRIPTION
    The overrides file maps a recipe id (the recipe's file name without the
    extension) to fields that replace the recipe's own when Yardstick runs it. It
    keeps site-specific values such as scope tags out of the shared recipe files.

    A missing or empty file means no overrides. Anything malformed throws rather
    than being skipped: an override that silently went missing would upload apps
    with, for example, the wrong scope tags.

    .PARAMETER Path
    The overrides file.

    .OUTPUTS
    A case-insensitive hashtable of recipe id to a hashtable of fields.
    #>
    param(
        [Parameter(Mandatory)][string]$Path
    )

    $overrides = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $overrides
    }

    $parsed = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Yaml
    if ($null -eq $parsed) {
        return $overrides
    }
    if ($parsed -isnot [System.Collections.IDictionary]) {
        throw "The file must map recipe ids to the fields they override."
    }

    foreach ($id in $parsed.Keys) {
        $fields = $parsed[$id]
        if ($fields -isnot [System.Collections.IDictionary]) {
            throw "The entry for '$id' must list the fields it overrides, such as scopeTags."
        }
        foreach ($reserved in 'id', 'base') {
            if ($fields.Keys -contains $reserved) {
                throw "The entry for '$id' cannot override '$reserved'."
            }
        }
        $overrides[[string]$id] = $fields
    }
    return $overrides
}


function Merge-RecipeWithOverride {
    <#
    .SYNOPSIS
    Applies a recipe's entry from the overrides file.

    .DESCRIPTION
    Each field in the override replaces the recipe's field whole: an overridden
    scopeTags list replaces the recipe's list rather than adding to it. Applied
    after Merge-RecipeWithBase, so an override also beats the base recipe.

    .PARAMETER Recipe
    The recipe hashtable, already merged with its base.

    .PARAMETER Override
    The recipe's entry from Import-RecipeOverrides.

    .OUTPUTS
    A merged hashtable.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Recipe,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Override
    )

    $merged = $Recipe.Clone()
    foreach ($key in $Override.Keys) {
        $merged[$key] = $Override[$key]
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

    # Runner-side scripts are dot-sourced into Yardstick's own scope (Invoke-Command
    # -NoNewScope), and PowerShell variable names are case-insensitive, so a recipe
    # that assigns an innocuous-looking $temp overwrites the runner's $Temp for the
    # rest of the run. Restore-RunnerPathVariable repairs that between recipes, but
    # the clobber still breaks the recipe doing it, so warn at authoring time.
    # Only these three fields: install/uninstall/detection scripts run on the client
    # and share nothing with the runner.
    $reservedVariablePattern = '(?im)^\s*\$(temp|buildspace|scripts|published|recipes|icons|tools|secrets|softwaredropbox|softwarearchive|prefs|applications)\s*='
    foreach ($field in 'preDownloadScript', 'downloadScript', 'postDownloadScript') {
        $body = [string]$Recipe[$field]
        if ([string]::IsNullOrWhiteSpace($body)) { continue }
        foreach ($match in [regex]::Matches($body, $reservedVariablePattern)) {
            $warnings.Add("Field '$field' assigns to reserved runner variable '`$$($match.Groups[1].Value)'. Rename it - it overwrites Yardstick's own folder variable for the rest of the run.")
        }
    }

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
    $detectionTypeFields = @{
        'file'     = @('fileDetectionPath', 'fileDetectionMethod', 'fileDetectionName')
        'registry' = @('registryDetectionMethod', 'registryDetectionKey')
        'script'   = @('detectionScript')
        'msi'      = @()
    }
    if ($Recipe.ContainsKey('detectionType') -and -not [string]::IsNullOrWhiteSpace($Recipe['detectionType'])) {
        $detType = $Recipe['detectionType'].ToString().ToLower()

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

    # additionalDetectionRules: extra file/registry rules that must all match
    # alongside the primary rule. Entries reuse the top-level field names. Intune
    # cannot combine a script rule with anything, and an msi entry would only
    # repeat the primary rule's product code, so both are rejected.
    if ($Recipe.ContainsKey('additionalDetectionRules') -and $null -ne $Recipe['additionalDetectionRules']) {
        $additionalRules = $Recipe['additionalDetectionRules']
        $ruleFields = @(
            'detectionType', 'fileDetectionMethod', 'fileDetectionPath', 'fileDetectionName',
            'fileDetectionOperator', 'fileDetectionVersion', 'fileDetectionDateTime', 'fileDetectionValue',
            'registryDetectionMethod', 'registryDetectionKey', 'registryDetectionValueName',
            'registryDetectionValue', 'registryDetectionOperator'
        )
        $ruleFieldsLower = $ruleFields | ForEach-Object { $_.ToLower() }

        if ($additionalRules -is [string] -or $additionalRules -is [System.Collections.IDictionary] -or $additionalRules -isnot [System.Collections.IEnumerable]) {
            $errors.Add("Field 'additionalDetectionRules' must be a list of detection rules")
        }
        else {
            if ([string]$Recipe['detectionType'] -eq 'script') {
                $errors.Add("Field 'additionalDetectionRules' cannot be used with detectionType 'script' - Intune does not allow a detection script to be combined with other rules")
            }
            $index = 0
            foreach ($rule in $additionalRules) {
                $label = "additionalDetectionRules[$index]"
                $index++
                if ($rule -isnot [System.Collections.IDictionary]) {
                    $errors.Add("Entry '$label' must be a mapping of detection fields")
                    continue
                }

                $ruleType = ([string]$rule['detectionType']).ToLower()
                if ($ruleType -notin @('file', 'registry')) {
                    $errors.Add("Invalid detectionType '$($rule['detectionType'])' in '$label'. Allowed values: file, registry")
                    continue
                }

                foreach ($field in 'fileDetectionMethod', 'registryDetectionMethod') {
                    if (-not [string]::IsNullOrWhiteSpace($rule[$field])) {
                        $allowed = $validEnumerations[$field]
                        if ($rule[$field].ToString().ToLower() -notin $allowed) {
                            $errors.Add("Invalid value '$($rule[$field])' for '$field' in '$label'. Allowed values: $($allowed -join ', ')")
                        }
                    }
                }

                # No script-assignment fallback here: scripts can only edit an entry
                # that already exists, so every required field must be in the YAML.
                foreach ($field in $detectionTypeFields[$ruleType]) {
                    if ([string]::IsNullOrWhiteSpace($rule[$field])) {
                        $errors.Add("Missing required field '$field' in '$label' (detectionType '$ruleType')")
                    }
                }
                $registryMethod = ([string]$rule['registryDetectionMethod']).ToLower()
                if ($ruleType -eq 'registry' -and $registryMethod -in @('string', 'integer', 'version')) {
                    $comparisonFields = @('registryDetectionValueName', 'registryDetectionOperator')
                    # version falls back to the recipe's version; string/integer have nothing to fall back to
                    if ($registryMethod -ne 'version') { $comparisonFields += 'registryDetectionValue' }
                    foreach ($field in $comparisonFields) {
                        if ([string]::IsNullOrWhiteSpace($rule[$field])) {
                            $errors.Add("Missing required field '$field' in '$label' (registryDetectionMethod '$registryMethod')")
                        }
                    }
                }

                foreach ($key in $rule.Keys) {
                    if ($key.ToLower() -notin $ruleFieldsLower) {
                        $warnings.Add("Unknown field '$key' in '$label' - this field is not used by Yardstick and may be a typo")
                    }
                    elseif ($key -cnotin $ruleFields) {
                        $expectedCasing = $ruleFields | Where-Object { $_.ToLower() -eq $key.ToLower() } | Select-Object -First 1
                        $warnings.Add("Field '$key' in '$label' has incorrect casing - expected '$expectedCasing'")
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
        'registryDetectionValue', 'registryDetectionOperator', 'additionalDetectionRules', 'detectionScript',
        'detectionScriptFileExtension', 'detectionScriptRunAs32Bit',
        'detectionScriptEnforceSignatureCheck', 'iconFile', 'description', 'publisher',
        'versionLock', 'numVersionsToKeep', 'fileType', 'softwareName',
        'manualDownload', 'manualDownloadFolder',
        'dependentApplicationBlocklist', 'dependentLinkUpdateEnabled',
        'dependentLinkUpdateRetryCount', 'dependentLinkUpdateRetryDelaySeconds',
        'dependentLinkUpdateTimeoutSeconds',
        'supersedence', 'uninstallPreviousVersion', 'autoUpdateOnAssignment', 'autoUpdate',
        'groupSkipAutoUpdates',
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
        Write-Log "Test-VersionExcluded was given an empty version." -Level Warning
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

    .PARAMETER MissingIcon
    The recipe's iconFile when it was not found in the Icons folder and the app was
    uploaded without an icon. Shown in the admin email. $null when the icon was used.
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
        [string]$AutoUpdateStatus,

        [AllowNull()]
        [string]$MissingIcon
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
        MissingIcon      = if ($MissingIcon) { $MissingIcon } else { $null }
        # Filled in later by Wait-YardstickBackup, once the background copy of
        # this app's .intunewin has finished.
        BackupStatus     = $null
        Timestamp        = Get-Date
    }
    
    $Script:SuccessfulApplications.Add($appInfo)
    # The log report reads the outcome from this line; keep its wording.
    Write-Log "Done: $DisplayName $Version ($Action)."
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

    # The single failure line for the app. The log report reads the stage and
    # message from it, so keep its wording.
    Write-Log "$FailureStage failed: $ErrorMessage" -Level Error

    # The email report lists each app once, under its first failure.
    if (-not ($Script:FailedApplications | Where-Object { $_.ApplicationId -eq $ApplicationId })) {
        $Script:FailedApplications.Add($appInfo)
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
        if ($outlook) { Write-Output "Works" | Out-Null } # This just gets rid of the IDE error since we're actually using the try catch block to test.
        $outlook = $null
        [System.GC]::Collect()
        return $true
    } catch {
        Write-Log "Outlook COM object not available: $_" -Level Detail
        return $false
    }
}


function Get-YardstickOutlookSecurityKey {
    <#
    .SYNOPSIS
    Returns the registry key that holds Outlook's programmatic access setting.

    .DESCRIPTION
    32-bit Office on 64-bit Windows reads the setting from the Wow6432Node key;
    otherwise it reads the native key. Office bitness comes from the Outlook
    Bitness value, falling back to the Click-to-Run Platform value.
    #>
    $bitness = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook' -Name 'Bitness' -ErrorAction SilentlyContinue).Bitness
    if (-not $bitness) {
        $bitness = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name 'Platform' -ErrorAction SilentlyContinue).Platform
    }
    if ($bitness -eq 'x86' -and [Environment]::Is64BitOperatingSystem) {
        return 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Outlook\Security'
    }
    return 'HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook\Security'
}


function Get-YardstickOutlookObjectModelGuard {
    <#
    .SYNOPSIS
    Reads Outlook's ObjectModelGuard setting.

    .OUTPUTS
    0 (warn when antivirus is inactive or out of date), 1 (always warn),
    2 (never warn), or $null when the value is not set.
    #>
    $value = (Get-ItemProperty -Path (Get-YardstickOutlookSecurityKey) -Name 'ObjectModelGuard' -ErrorAction SilentlyContinue).ObjectModelGuard
    if ($null -eq $value) { return $null }
    return [int]$value
}


function Set-YardstickOutlookObjectModelGuard {
    <#
    .SYNOPSIS
    Turns off Outlook's programmatic access warnings for this computer.

    .DESCRIPTION
    Sets ObjectModelGuard to 2 ("Never warn me about suspicious activity"), so
    Outlook stops asking for approval when Yardstick sends an email. The setting
    applies to every program on the computer and needs an elevated session.
    #>
    $principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Turning off Outlook's security prompts needs an elevated PowerShell (Run as administrator)."
    }
    $key = Get-YardstickOutlookSecurityKey
    if (-not (Test-Path -Path $key)) {
        New-Item -Path $key -Force | Out-Null
    }
    New-ItemProperty -Path $key -Name 'ObjectModelGuard' -PropertyType DWord -Value 2 -Force | Out-Null
}


function Test-YardstickOutlookPromptExpected {
    <#
    .SYNOPSIS
    Tests whether Outlook will ask for approval each time Yardstick sends an email.

    .DESCRIPTION
    With the default setting, Outlook prompts only when Windows Security Center
    reports the antivirus as inactive or out of date. Windows Server has no
    Security Center, so there Outlook always prompts unless ObjectModelGuard is 2.
    #>
    $guard = Get-YardstickOutlookObjectModelGuard
    if ($guard -eq 2) { return $false }
    if ($guard -eq 1) { return $true }
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
    return [bool]($os -and $os.ProductType -ne 1)
}


function Get-YardstickEmailDeliveryMethod {
    <#
    .SYNOPSIS
    Resolves the configured email delivery method.

    .OUTPUTS
    The lowercased emailDeliveryMethod preference, or 'outlook' when it is unset.
    #>
    param($Preferences)

    if (-not $Preferences -or [string]::IsNullOrWhiteSpace([string]$Preferences.emailDeliveryMethod)) {
        return 'outlook'
    }
    return ([string]$Preferences.emailDeliveryMethod).Trim().ToLowerInvariant()
}


function Send-YardstickSmtpMessage {
    <#
    .SYNOPSIS
    Sends an HTML email through an SMTP server using Send-MailMessage.

    .DESCRIPTION
    Reads smtpServer, smtpPort (default 25) and smtpUseSsl from preferences, and
    sends from emailSendFromAddress, displayed as emailSenderName when set. When an
    SMTP credential is stored in Windows Credential Manager (see
    Set-YardstickCredential.ps1 -Smtp) it is used to authenticate; otherwise the
    message is relayed anonymously. Errors are thrown to the caller.

    .PARAMETER Preferences
    Hashtable containing SMTP and email configuration preferences.

    .PARAMETER To
    One or more recipient addresses.

    .PARAMETER Subject
    Message subject.

    .PARAMETER HtmlBody
    HTML message body.
    #>
    param(
        [Parameter(Mandatory=$true)]
        $Preferences,

        [Parameter(Mandatory=$true)]
        [string[]]$To,

        [Parameter(Mandatory=$true)]
        [string]$Subject,

        [Parameter(Mandatory=$true)]
        [string]$HtmlBody
    )

    if (-not $Preferences.smtpServer) {
        throw "smtpServer is not configured in preferences."
    }
    if (-not $Preferences.emailSendFromAddress) {
        throw "emailSendFromAddress is not configured in preferences."
    }

    $from = if ($Preferences.emailSenderName) {
        "$($Preferences.emailSenderName) <$($Preferences.emailSendFromAddress)>"
    } else {
        [string]$Preferences.emailSendFromAddress
    }
    $port = if (($Preferences.smtpPort -as [int]) -gt 0) { [int]$Preferences.smtpPort } else { 25 }

    $mailParams = @{
        SmtpServer = [string]$Preferences.smtpServer
        Port       = $port
        From       = $from
        To         = $To
        Subject    = $Subject
        Body       = $HtmlBody
        BodyAsHtml = $true
        Encoding   = [System.Text.Encoding]::UTF8
        UseSsl     = [bool]$Preferences.smtpUseSsl
    }

    # The SMTP credential lives in YardstickCredential.psm1, which is not always loaded.
    if (Get-Command -Name Get-YardstickSmtpCredential -ErrorAction SilentlyContinue) {
        $credential = Get-YardstickSmtpCredential -Target (Get-YardstickSmtpCredentialTarget -Preferences $Preferences)
        if ($credential) {
            $mailParams.Credential = $credential
        }
    }

    # Send-MailMessage is marked obsolete in PowerShell 7 and warns on every call.
    Send-MailMessage @mailParams -ErrorAction Stop -WarningAction SilentlyContinue
}


function Get-YardstickEmailAudiences {
    <#
    .SYNOPSIS
    Lists the report audiences a run should email.

    .DESCRIPTION
    The admin report (everything, including failures) is always sent. It goes to
    adminEmailRecipient, or to emailRecipient when no admin address is configured.
    The general report (updated software only) is added when adminEmailRecipient
    and emailRecipient are both configured, so the wider list is not sent the
    admin report as well.

    .OUTPUTS
    'Admin' and, when applicable, 'General'.
    #>
    param([Parameter(Mandatory=$true)]$Preferences)

    $audiences = @('Admin')
    if ($Preferences['adminEmailRecipient'] -and $Preferences['emailRecipient']) {
        $audiences += 'General'
    }
    return $audiences
}


function Send-YardstickEmailReport {
    <#
    .SYNOPSIS
    Sends the email reports of application processing results.

    .DESCRIPTION
    Sends two emails per run. The admin email carries everything: the run
    parameters, successful and failed applications, dependents, auto-update and
    backup status. The general email lists only the software that was updated.
    With no adminEmailRecipient configured, a single admin email goes to
    emailRecipient. The general email is skipped when nothing was updated.
    Delivery follows emailDeliveryMethod: Microsoft Outlook's COM interface by
    default, or Send-MailMessage when set to 'smtp'.

    .PARAMETER Preferences
    Hashtable containing email configuration preferences.

    .PARAMETER RunParameters
    String describing the parameters used for this Yardstick run.

    .PARAMETER Audience
    'All' (default) sends every applicable report. 'Admin' or 'General' sends just that one.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Preferences,

        [string]$RunParameters = "",

        [ValidateSet('All', 'Admin', 'General')]
        [string]$Audience = 'All',

        # When set, the email is opened for visual inspection instead of being sent -
        # in Outlook, or in the default browser for SMTP delivery.
        [switch]$Preview,

        # Optional path. If provided, the rendered HTML body is also written to this file
        # so it can be opened in a browser (useful when validating formatting changes).
        # The general report is written alongside it with a ".general" suffix.
        [string]$HtmlOutputPath
    )
    
    # Check if email notifications are enabled
    if (-not $Preferences.emailNotificationEnabled) {
        Write-Log "Email notifications are disabled in preferences."
        return
    }

    if ($Audience -eq 'All') {
        foreach ($target in (Get-YardstickEmailAudiences -Preferences $Preferences)) {
            $audienceArgs = @{}
            foreach ($key in $PSBoundParameters.Keys) { $audienceArgs[$key] = $PSBoundParameters[$key] }
            $audienceArgs.Audience = $target
            if ($HtmlOutputPath -and $target -eq 'General') {
                $audienceArgs.HtmlOutputPath = [System.IO.Path]::ChangeExtension($HtmlOutputPath, ".general$([System.IO.Path]::GetExtension($HtmlOutputPath))")
            }
            Send-YardstickEmailReport @audienceArgs
        }
        return
    }

    $isAdmin = $Audience -eq 'Admin'
    $separateAdmin = [bool]$Preferences['adminEmailRecipient']
    $recipients = if ($isAdmin -and $separateAdmin) { @($Preferences['adminEmailRecipient']) } else { @($Preferences['emailRecipient']) }
    $subject = [string]$Preferences['emailSubject']
    if ($isAdmin -and $separateAdmin) {
        $subject = if ($Preferences['adminEmailSubject']) { [string]$Preferences['adminEmailSubject'] } else { "$subject (Admin)" }
    }

    # Validate required email settings
    if (-not $recipients) {
        Write-Log "No recipients are configured for the $Audience email report. Skipping it." -Level Warning
        return
    }
    foreach ($setting in @('emailSubject', 'emailSenderName')) {
        if (-not $Preferences[$setting]) {
            Write-Log "Email setting '$setting' is not configured. Skipping the email report." -Level Warning
            return
        }
    }

    # Initialize tracking arrays if they don't exist
    if (-not $Script:SuccessfulApplications) {
        $Script:SuccessfulApplications = [System.Collections.Generic.List[PSObject]]::new()
    }
    if (-not $Script:FailedApplications) {
        $Script:FailedApplications = [System.Collections.Generic.List[PSObject]]::new()
    }

    # Maintenance sweeps and repairs republish nothing new, so they are not
    # "updated software" as far as the general audience is concerned.
    $updatedApplications = @($Script:SuccessfulApplications | Where-Object { $_.Action -in @('Updated', 'Force Updated') })
    if (-not $isAdmin -and $updatedApplications.Count -eq 0) {
        Write-Log "No applications were updated; skipping the general email report."
        return
    }
    
    $method = Get-YardstickEmailDeliveryMethod -Preferences $Preferences

    if ($method -eq 'smtp') {
        foreach ($setting in @('smtpServer', 'emailSendFromAddress')) {
            if (-not $Preferences[$setting]) {
                Write-Log "Email setting '$setting' is required for SMTP delivery. Skipping the email report." -Level Warning
                return
            }
        }
    } elseif (-not (Test-OutlookAvailability)) {
        Write-Log "Outlook is not available. Cannot send the email report." -Level Warning
        return
    }

    try {
        Write-Log "Creating the $Audience email report..." -Level Detail

        # Load branding logo. Use a cid: reference for the email itself (Outlook
        # desktop does not render base64 data URIs reliably) and a base64 data URI
        # for the HtmlOutputPath so the standalone file stays self-contained.
        # The cid is the attachment's file name, which Outlook resolves on its own;
        # setting a Content-ID would need Attachment.PropertyAccessor, which raises
        # an Outlook security prompt.
        # Send-MailMessage cannot set a Content-ID on an attachment, and many
        # clients block data URIs, so SMTP reports go out without the logo.
        $logoHtml = ""
        $logoHtmlForBrowser = ""
        $logoPath = $null
        if ($method -ne 'smtp') {
            try {
                $candidatePath = Join-Path $PSScriptRoot "..\Branding\yardstick_logo_white_text_transparent_bg.png"
                if (Test-Path $candidatePath) {
                    $logoPath = (Resolve-Path $candidatePath).ProviderPath
                    $logoCid = [System.IO.Path]::GetFileName($logoPath)
                    $logoBytes = [System.IO.File]::ReadAllBytes($logoPath)
                    $logoBase64 = [System.Convert]::ToBase64String($logoBytes)
                    $logoHtml = "<img src=`"cid:$logoCid`" alt=`"Yardstick`" width=`"180`" height=`"180`" class=`"header-logo`" style=`"width:180px; height:180px; max-width:100%; display:block; border:0;`" />"
                    $logoHtmlForBrowser = "<img src=`"data:image/png;base64,$logoBase64`" alt=`"Yardstick`" width=`"180`" height=`"180`" class=`"header-logo`" style=`"width:180px; height:180px; max-width:100%; display:block; border:0;`" />"
                } else {
                    Write-Log "Branding logo not found at $candidatePath." -Level Warning
                }
            } catch {
                Write-Log "Could not embed the branding logo: $($_.Exception.Message)" -Level Warning
            }
        }

        # Without a logo, drop its cell rather than leave an empty 195px gap.
        $logoCell = if ($logoHtml) { "<td class=`"header-logo-cell`" style=`"width:195px; padding:15px 15px 15px 15px; vertical-align:middle; border:0;`">$logoHtml</td>" } else { "" }
        $headerTextPadding = if ($logoHtml) { "15px 15px 15px 0" } else { "15px" }

        $summaryRowStyle = 'margin:0; padding:0; line-height:1.3; color:#1a1a1a;'
        $summaryRows = if ($isAdmin) {
            @"
                    <p style="$summaryRowStyle"><strong>Successful Applications:</strong> $($Script:SuccessfulApplications.Count)</p>
                    <p style="$summaryRowStyle"><strong>Failed Applications:</strong> $($Script:FailedApplications.Count)</p>
                    <p style="$summaryRowStyle"><strong>Total Processed:</strong> $($Script:SuccessfulApplications.Count + $Script:FailedApplications.Count)</p>
"@
        } else {
            @"
                    <p style="$summaryRowStyle"><strong>Updated Applications:</strong> $($updatedApplications.Count)</p>
"@
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
            $logoCell
            <td class="header-text-cell" style="padding:$headerTextPadding; vertical-align:middle; border:0;">
                <p style="margin:0 0 6px 0; padding:0; font-size:18pt; font-weight:bold; line-height:1.1; color:white;">Application Update Report</p>
                <p style="margin:0; padding:0; line-height:1.2; color:white;">Run Time: $(Get-Date -Format "MMMM dd, yyyy 'at' HH:mm:ss tt")</p>
$(if ($isAdmin -and $RunParameters) { @"
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
$summaryRows
                </td>
            </tr>
        </table>
"@

        if ($isAdmin -and $Script:SuccessfulApplications.Count -gt 0) {
            # Only show the Backup column when backups are actually configured,
            # so runs without a Backup preference do not carry an empty column.
            $anyBackup = @($Script:SuccessfulApplications | Where-Object { $_.PSObject.Properties['BackupStatus'] -and $_.BackupStatus }).Count -gt 0
            $backupHeader = if ($anyBackup) { "`n                            <th>Backup</th>" } else { "" }
            $successColSpan = if ($anyBackup) { 6 } else { 5 }
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
                            <th>Auto-Update</th>$backupHeader
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
                $backupCell = ""
                if ($anyBackup) {
                    $bkStatus = if ($app.PSObject.Properties['BackupStatus'] -and $app.BackupStatus) { $app.BackupStatus } else { '&mdash;' }
                    $bkColor = switch -Regex ($bkStatus) {
                        '^ok'              { '#28a745' }
                        '^skipped'         { '#6c757d' }
                        '^failed|^timed'   { '#dc3545' }
                        default            { '#6c757d' }
                    }
                    $backupCell = "`n                    <td data-label=`"Backup`"><span style=`"color: $bkColor;`">$bkStatus</span></td>"
                }
                # Uploaded without an icon: flag it so the operator knows to add one
                # (Get-YardstickIcon.ps1 -RecipeId <id>).
                $iconNote = ""
                if ($app.PSObject.Properties['MissingIcon'] -and $app.MissingIcon) {
                    $iconNote = "<br><small style=`"color:#b35c00;`">No icon uploaded: $([System.Net.WebUtility]::HtmlEncode([string]$app.MissingIcon)) is missing from the Icons folder</small>"
                }
                $emailBody += @"
                <tr>
                    <td data-label="Application"><strong>$($app.DisplayName)</strong><br><small>ID: $($app.ApplicationId)</small>$iconNote</td>
                    <td data-label="Version">$($app.Version)</td>
                    <td data-label="Action">$($app.Action)</td>
                    <td data-label="Auto-Update"><span style="color: $auColor;">$auStatus</span></td>$backupCell
                    <td data-label="Time" class="timestamp">$($app.Timestamp.ToString("MM/dd/yyyy HH:mm:ss"))</td>
                </tr>
"@
                # Add dependency information if present
                if ($app.Dependents -and $app.Dependents.Count -gt 0) {
                    $emailBody += @"
                <tr class="dependents-row">
                    <td colspan="$successColSpan" style="background-color: #f8f9fa; color:#1a1a1a; padding-left: 30px;">
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
        if ($isAdmin -and $Script:FailedApplications.Count -gt 0) {
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

        if (-not $isAdmin) {
            $emailBody += @"

        <table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#d4edda" style="width:100%; background-color:#d4edda; border-collapse:separate; border-radius:5px; margin:10px 0;">
            <tr>
                <td class="section-success" bgcolor="#d4edda" style="background-color:#d4edda; color:#1a1a1a; padding:15px; border-left:4px solid #28a745; border-radius:5px;">
                    <h3 style="margin:0 0 8px 0; padding:0; line-height:1.2; color:#1a1a1a;">Updated Applications</h3>
                    <table class="report-table">
                        <tr class="report-header-row">
                            <th>Application</th>
                            <th>Version</th>
                            <th>Time</th>
                        </tr>
"@
            foreach ($app in $updatedApplications) {
                $encodedName = [System.Net.WebUtility]::HtmlEncode([string]$app.DisplayName)
                $encodedVersion = [System.Net.WebUtility]::HtmlEncode([string]$app.Version)
                $emailBody += @"
                <tr>
                    <td data-label="Application"><strong>$encodedName</strong></td>
                    <td data-label="Version">$encodedVersion</td>
                    <td data-label="Time" class="timestamp">$($app.Timestamp.ToString("MM/dd/yyyy HH:mm:ss"))</td>
                </tr>
"@
            }
            $emailBody += @"
                    </table>
                </td>
            </tr>
        </table>
"@
        }

        # Add footer
        $footerLogLine = if ($isAdmin) { "<br>`n                Log file location: $Global:LogLocation\$Global:LogFile" } else { "" }
        $emailBody += @"
        
        <hr style="margin: 20px 0;">
        <p class="timestamp">
            <small>
                This report was automatically generated by Yardstick.$footerLogLine
            </small>
        </p>
    </div>
</body>
</html>
"@

        if ($HtmlOutputPath) {
            try {
                $htmlForFile = $emailBody
                if ($logoHtmlForBrowser -and $logoHtml) {
                    $htmlForFile = $htmlForFile.Replace($logoHtml, $logoHtmlForBrowser)
                }
                Set-Content -Path $HtmlOutputPath -Value $htmlForFile -Encoding UTF8
                Write-Log "Wrote the rendered email HTML to $HtmlOutputPath."
            } catch {
                Write-Log "Could not write the rendered HTML to ${HtmlOutputPath}: $_" -Level Warning
            }
        }

        if ($method -eq 'smtp') {
            if ($Preview) {
                # SMTP has no compose window to display, so preview in the browser.
                $previewPath = $HtmlOutputPath
                if (-not $previewPath) {
                    $previewPath = Join-Path ([System.IO.Path]::GetTempPath()) "YardstickEmailPreview.$Audience.html"
                    Set-Content -Path $previewPath -Value $emailBody -Encoding UTF8
                }
                Invoke-Item -LiteralPath $previewPath
                Write-Log "Email opened in the default browser for preview (not sent): $previewPath"
            } else {
                Send-YardstickSmtpMessage -Preferences $Preferences -To $recipients `
                    -Subject $subject -HtmlBody $emailBody
                Write-Log "$Audience email report sent via SMTP to: $($recipients -join ', ')."
            }
            return
        }

        # Create Outlook application and mail item
        $outlook = New-Object -ComObject "Outlook.Application"
        $mail = $outlook.CreateItem(0)  # 0 = olMailItem

        # Set email properties
        $mail.To = ($recipients -join "; ")

        $mail.Subject = $subject
        if ($Preferences.emailSendFromAddress) {
            $mail.SentOnBehalfOfName = $Preferences.emailSendFromAddress
        }

        # Attach the logo before setting the body, so Outlook matches the
        # cid:<file name> reference to it and shows it inline.
        if ($logoPath) {
            try {
                $mail.Attachments.Add($logoPath) | Out-Null
            } catch {
                Write-Log "Could not attach the logo: $($_.Exception.Message)" -Level Warning
            }
        }

        # Set email body and send (or open for preview)
        $mail.HTMLBody = $emailBody

        if ($Preview) {
            $mail.Display()
            Write-Log "Email opened in Outlook for preview (not sent)."
        } else {
            $mail.Send()
            Write-Log "$Audience email report sent to: $($recipients -join ', ')."
        }

        # Clean up COM objects
        $mail = $null
        $outlook = $null
        [System.GC]::Collect()
    } catch {
        Write-Log "Could not send the email report: $_" -Level Error
        # Clean up COM objects even on error
        try {
            $mail = $null
            $outlook = $null
            [System.GC]::Collect()
        } catch { }
    }
}


function Add-NativeBitnessGuard {
    <#
    .SYNOPSIS
    Prepends a 32-bit re-launch guard to a generated install.ps1/uninstall.ps1.

    .DESCRIPTION
    The Intune Management Extension is a 32-bit process, and the command lines
    Yardstick auto-generates for powerShellInstallScript/powerShellUninstallScript
    invoke a bare "powershell.exe". A bare name resolves by PATH, which under
    WOW64 means SysWOW64 - so the script runs 32-bit and silently sees a
    redirected machine: HKLM:\SOFTWARE\... reads come from WOW6432Node, and a
    64-bit application's ARP key and install directory are simply invisible. An
    uninstaller written against those paths then finds nothing, removes nothing,
    and exits 0, so Intune reports success while the app is still installed.

    The guard re-launches the script once under native PowerShell and propagates
    the child's exit code. It is deliberately a script preamble rather than a
    "%SystemRoot%\sysnative\..." command line: sysnative is a virtual alias that
    exists only inside a 32-bit process, so a sysnative command line fails
    outright whenever the caller is already 64-bit. The preamble is correct in
    both cases, which keeps these packages testable by the recipe harness and
    keeps them working if the Management Extension ever ships 64-bit.

    .PARAMETER ScriptContent
    The recipe-authored script body, after placeholder replacement.

    .OUTPUTS
    The script body with the guard prepended.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string] $ScriptContent
    )

    # -WorkingDirectory is passed explicitly because Start-Process otherwise
    # resolves against the PowerShell provider location, which is not reliably
    # the package directory the cmd wrapper cd'd into. Install scripts routinely
    # reference the payload by a relative path, so this has to survive.
    $guard = @'
# --- Begin Yardstick native-bitness guard (generated - do not edit) ---
# The Intune Management Extension is 32-bit, so the generated command line
# starts this script under WOW64: HKLM:\SOFTWARE reads are redirected to
# WOW6432Node and 64-bit install paths are hidden. Re-launch once natively so
# the script sees the real machine.
if (-not [Environment]::Is64BitProcess -and [Environment]::Is64BitOperatingSystem) {
    $yardstickNativePowerShell = Join-Path $env:SystemRoot 'sysnative\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $yardstickNativePowerShell) {
        $yardstickChild = Start-Process -FilePath $yardstickNativePowerShell -Wait -PassThru -NoNewWindow `
            -WorkingDirectory (Get-Location -PSProvider FileSystem).ProviderPath `
            -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`""
        exit $yardstickChild.ExitCode
    }
    Write-Output 'WARNING: running 32-bit and native PowerShell was not found; continuing under WOW64 redirection.'
}
# --- End Yardstick native-bitness guard ---

'@

    return $guard + $ScriptContent
}


# Export all functions for module availability
Export-ModuleMember -Function *
