<#
.SYNOPSIS
Helper functions for the local recipe install/uninstall test harness.

.DESCRIPTION
These functions evaluate Intune-style detection rules locally, run installer
command lines with timeouts and full output capture, and snapshot machine state
so that leftovers from an uninstall can be reported.

Every function here is self-contained. Nothing in this module reads variables
from the caller's scope, which keeps it safe to import from any harness script.
#>

Set-StrictMode -Version Latest

$script:UninstallKeyPaths = @(
    @{ Hive = 'LocalMachine'; View = 'Registry64'; Path = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' }
    @{ Hive = 'LocalMachine'; View = 'Registry32'; Path = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' }
    @{ Hive = 'CurrentUser';  View = 'Registry64'; Path = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' }
)

function Test-IsElevated {
    [CmdletBinding()]
    param()
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    return (New-Object Security.Principal.WindowsPrincipal $identity).IsInRole(
        [Security.Principal.WindowsBuiltinRole]::Administrator)
}

function Expand-RecipeString {
    [CmdletBinding()]
    param([AllowNull()][string] $Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Value }
    return [Environment]::ExpandEnvironmentVariables($Value.Trim("'").Trim('"'))
}

function Expand-RecipeToken {
    <#
    .SYNOPSIS
    Replaces the Yardstick <filename>, <version>, and <productcode> placeholders.
    Mirrors Set-ScriptPlaceholders in Yardstick.ps1.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][string] $Value,
        [string] $FileName,
        [string] $Version,
        [AllowNull()][string] $ProductCode
    )
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Value }
    $result = $Value
    if ($FileName) { $result = $result.Replace('<filename>', $FileName) }
    if ($Version) { $result = $result.Replace('<version>', $Version) }
    if ($ProductCode) { $result = $result.Replace('<productcode>', $ProductCode) }
    return $result
}

function ConvertTo-ComparableVersion {
    <#
    .SYNOPSIS
    Best-effort conversion of a version string to [version] for comparison.
    Falls back to string comparison data when the value is not version shaped.
    #>
    [CmdletBinding()]
    param([AllowNull()][string] $Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $candidate = ($Value.Trim() -replace '[^0-9.].*$', '').Trim('.')
    if ([string]::IsNullOrWhiteSpace($candidate)) { return $null }
    $parts = @($candidate -split '\.' | Where-Object { $_ -ne '' } | Select-Object -First 4)
    if ($parts.Count -eq 0) { return $null }
    if ($parts.Count -eq 1) { $parts += '0' }
    $normalized = ($parts | ForEach-Object { [int]$_ }) -join '.'
    $parsed = $null
    if ([version]::TryParse($normalized, [ref]$parsed)) { return $parsed }
    return $null
}

function Compare-DetectionValue {
    <#
    .SYNOPSIS
    Applies an Intune detection operator to two values.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()] $Actual,
        [Parameter(Mandatory)][AllowNull()] $Expected,
        [string] $Operator = 'greaterThanOrEqual',
        [ValidateSet('version', 'integer', 'string', 'datetime')][string] $Kind = 'version'
    )

    if ($null -eq $Actual) { return $false }

    switch ($Kind) {
        'version' {
            $left = ConvertTo-ComparableVersion ([string]$Actual)
            $right = ConvertTo-ComparableVersion ([string]$Expected)
            if ($null -eq $left -or $null -eq $right) {
                # Not version shaped on at least one side; only equality is meaningful.
                $equal = ([string]$Actual).Trim() -eq ([string]$Expected).Trim()
                return $(if ($Operator -eq 'notEqual') { -not $equal } else { $equal })
            }
            $comparison = $left.CompareTo($right)
        }
        'integer' {
            $left = 0; $right = 0
            if (-not [long]::TryParse([string]$Actual, [ref]$left)) { return $false }
            if (-not [long]::TryParse([string]$Expected, [ref]$right)) { return $false }
            $comparison = $left.CompareTo($right)
        }
        'datetime' {
            $left = [datetime]::MinValue; $right = [datetime]::MinValue
            if (-not [datetime]::TryParse([string]$Actual, [ref]$left)) { return $false }
            if (-not [datetime]::TryParse([string]$Expected, [ref]$right)) { return $false }
            $comparison = $left.CompareTo($right)
        }
        default {
            $comparison = [string]::Compare([string]$Actual, [string]$Expected, $true)
        }
    }

    switch ($Operator) {
        'equal' { return $comparison -eq 0 }
        'notEqual' { return $comparison -ne 0 }
        'greaterThan' { return $comparison -gt 0 }
        'greaterThanOrEqual' { return $comparison -ge 0 }
        'lessThan' { return $comparison -lt 0 }
        'lessThanOrEqual' { return $comparison -le 0 }
        default { return $comparison -ge 0 }
    }
}

function Get-InstalledPackage {
    <#
    .SYNOPSIS
    Enumerates Add/Remove Programs entries from HKLM (64 and 32 bit) and HKCU.
    #>
    [CmdletBinding()]
    param()

    $packages = [Collections.Generic.List[psobject]]::new()
    foreach ($location in $script:UninstallKeyPaths) {
        try {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($location.Hive, $location.View)
            $root = $base.OpenSubKey($location.Path)
            if (-not $root) { continue }
            foreach ($name in $root.GetSubKeyNames()) {
                $key = $root.OpenSubKey($name)
                if (-not $key) { continue }
                $displayName = [string]$key.GetValue('DisplayName')
                if ([string]::IsNullOrWhiteSpace($displayName)) { continue }
                $packages.Add([PSCustomObject]@{
                        KeyId           = "$($location.Hive)/$($location.View)/$name"
                        SubKeyName      = $name
                        DisplayName     = $displayName
                        DisplayVersion  = [string]$key.GetValue('DisplayVersion')
                        Publisher       = [string]$key.GetValue('Publisher')
                        InstallLocation = [string]$key.GetValue('InstallLocation')
                        UninstallString = [string]$key.GetValue('UninstallString')
                        Scope           = $location.Hive
                        View            = $location.View
                    })
                $key.Dispose()
            }
            $root.Dispose()
            $base.Dispose()
        } catch {
            Write-Verbose "Could not read $($location.Hive)/$($location.View): $_"
        }
    }
    return $packages
}

function Get-SystemSnapshot {
    <#
    .SYNOPSIS
    Captures the machine state that an install is expected to change.
    #>
    [CmdletBinding()]
    param()

    $directoryRoots = @(
        $env:ProgramFiles
        ${env:ProgramFiles(x86)}
        (Join-Path $env:LOCALAPPDATA 'Programs')
        (Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs')
        (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique

    $directories = [Collections.Generic.List[string]]::new()
    foreach ($root in $directoryRoots) {
        try {
            foreach ($child in [IO.Directory]::EnumerateFileSystemEntries($root)) {
                $directories.Add($child)
            }
        } catch {
            Write-Verbose "Could not enumerate '$root': $_"
        }
    }

    [PSCustomObject]@{
        TakenAt     = (Get-Date).ToString('o')
        Packages    = @(Get-InstalledPackage)
        Directories = @($directories)
    }
}

function Compare-SystemSnapshot {
    <#
    .SYNOPSIS
    Reports what a snapshot pair says was added or removed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Before,
        [Parameter(Mandatory)] $After
    )

    $beforeKeys = @($Before.Packages | ForEach-Object { $_.KeyId })
    $afterKeys = @($After.Packages | ForEach-Object { $_.KeyId })
    $beforeDirs = @($Before.Directories)
    $afterDirs = @($After.Directories)

    [PSCustomObject]@{
        AddedPackages     = @($After.Packages | Where-Object { $_.KeyId -notin $beforeKeys })
        RemovedPackages   = @($Before.Packages | Where-Object { $_.KeyId -notin $afterKeys })
        AddedDirectories  = @($afterDirs | Where-Object { $_ -notin $beforeDirs })
        RemovedDirectories = @($beforeDirs | Where-Object { $_ -notin $afterDirs })
    }
}

function Invoke-TestProcess {
    <#
    .SYNOPSIS
    Runs a process with a timeout and captures stdout, stderr, and exit code.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $FilePath,
        [string[]] $ArgumentList = @(),
        [string] $WorkingDirectory = (Get-Location).Path,
        [int] $TimeoutSeconds = 1800,
        [Parameter(Mandatory)][string] $LogDirectory,
        [Parameter(Mandatory)][string] $LogName
    )

    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    $stdoutPath = Join-Path $LogDirectory "$LogName.out.log"
    $stderrPath = Join-Path $LogDirectory "$LogName.err.log"

    $startParams = @{
        FilePath               = $FilePath
        WorkingDirectory       = $WorkingDirectory
        RedirectStandardOutput = $stdoutPath
        RedirectStandardError  = $stderrPath
        PassThru               = $true
        WindowStyle            = 'Hidden'
    }
    if ($ArgumentList.Count -gt 0) { $startParams['ArgumentList'] = $ArgumentList }

    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    $timedOut = $false

    # Intune runs install command lines from the package directory, where a bare
    # "setup.exe" resolves. Some shells set NoDefaultCurrentDirectoryInExePath,
    # which breaks that resolution, so it is cleared for the child process only.
    $noCurrentDir = $env:NoDefaultCurrentDirectoryInExePath
    if ($null -ne $noCurrentDir) { Remove-Item Env:\NoDefaultCurrentDirectoryInExePath -ErrorAction SilentlyContinue }
    try {
        $process = Start-Process @startParams
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            $timedOut = $true
            try {
                Stop-Process -Id $process.Id -Force -ErrorAction Stop
                $process.WaitForExit(30000) | Out-Null
            } catch {
                Write-Warning "Could not stop timed-out process $($process.Id): $_"
            }
        }
    } finally {
        if ($null -ne $noCurrentDir) { $env:NoDefaultCurrentDirectoryInExePath = $noCurrentDir }
        $stopwatch.Stop()
    }

    $exitCode = if ($timedOut) { $null } else { $process.ExitCode }
    $stdout = if (Test-Path -LiteralPath $stdoutPath) { (Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue) } else { '' }
    $stderr = if (Test-Path -LiteralPath $stderrPath) { (Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue) } else { '' }

    [PSCustomObject]@{
        FilePath    = $FilePath
        Arguments   = ($ArgumentList -join ' ')
        ExitCode    = $exitCode
        TimedOut    = $timedOut
        DurationSec = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
        StdOut      = [string]$stdout
        StdErr      = [string]$stderr
        StdOutPath  = $stdoutPath
        StdErrPath  = $stderrPath
    }
}

function Invoke-CommandLine {
    <#
    .SYNOPSIS
    Runs a Yardstick install/uninstall command line the way Intune does, through
    cmd.exe, so that quoting and %ENVVAR% expansion behave identically.

    .DESCRIPTION
    The command line is written to a generated .cmd file rather than passed as
    cmd.exe arguments. PowerShell re-quotes arguments that contain embedded
    double quotes, which corrupts command lines such as
    "%LocalAppData%\...\Uninstall App.exe" /S. A batch file receives the text
    verbatim and still expands %ENVVAR% references.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $CommandLine,
        [Parameter(Mandatory)][string] $WorkingDirectory,
        [int] $TimeoutSeconds = 1800,
        [Parameter(Mandatory)][string] $LogDirectory,
        [Parameter(Mandatory)][string] $LogName
    )

    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    $batchPath = Join-Path $LogDirectory "$LogName.cmd"
    $batchLines = @(
        '@echo off'
        "cd /d `"$WorkingDirectory`""
        $CommandLine
        'exit /b %ERRORLEVEL%'
    )
    Set-Content -LiteralPath $batchPath -Value $batchLines -Encoding ASCII

    $run = Invoke-TestProcess -FilePath "$env:SystemRoot\System32\cmd.exe" `
        -ArgumentList @('/d', '/c', $batchPath) `
        -WorkingDirectory $WorkingDirectory `
        -TimeoutSeconds $TimeoutSeconds `
        -LogDirectory $LogDirectory `
        -LogName $LogName

    $run | Add-Member -NotePropertyName CommandLine -NotePropertyValue $CommandLine -Force
    $run | Add-Member -NotePropertyName BatchFile -NotePropertyValue $batchPath -Force
    return $run
}

function Invoke-DetectionScript {
    <#
    .SYNOPSIS
    Runs a detection script under Windows PowerShell and applies Intune rules:
    detected means exit code 0 *and* non-empty STDOUT.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $ScriptContent,
        [Parameter(Mandatory)][string] $LogDirectory,
        [Parameter(Mandatory)][string] $LogName,
        [switch] $RunAs32Bit,
        [int] $TimeoutSeconds = 300
    )

    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    $scriptPath = Join-Path $LogDirectory "$LogName.ps1"
    Set-Content -LiteralPath $scriptPath -Value $ScriptContent -Encoding UTF8

    $powerShell = if ($RunAs32Bit) {
        "$env:SystemRoot\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    } else {
        "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    }
    if (-not (Test-Path -LiteralPath $powerShell)) {
        $powerShell = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    }

    $run = Invoke-TestProcess -FilePath $powerShell `
        -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath) `
        -WorkingDirectory $LogDirectory `
        -TimeoutSeconds $TimeoutSeconds `
        -LogDirectory $LogDirectory `
        -LogName $LogName

    $detected = ($run.ExitCode -eq 0) -and (-not [string]::IsNullOrWhiteSpace($run.StdOut))
    [PSCustomObject]@{
        Detected = $detected
        Detail   = "exit=$($run.ExitCode) stdout='$(($run.StdOut -replace '\s+', ' ').Trim())'"
        Run      = $run
    }
}

function Get-RegistryValueForDetection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $KeyPath,
        [AllowNull()][string] $ValueName,
        [switch] $Check32BitOn64System
    )

    $normalized = $KeyPath.Trim().TrimEnd('\')
    $hiveMap = @{
        'HKEY_LOCAL_MACHINE'  = [Microsoft.Win32.RegistryHive]::LocalMachine
        'HKLM'                = [Microsoft.Win32.RegistryHive]::LocalMachine
        'HKEY_CURRENT_USER'   = [Microsoft.Win32.RegistryHive]::CurrentUser
        'HKCU'                = [Microsoft.Win32.RegistryHive]::CurrentUser
        'HKEY_CLASSES_ROOT'   = [Microsoft.Win32.RegistryHive]::ClassesRoot
        'HKEY_USERS'          = [Microsoft.Win32.RegistryHive]::Users
        'HKEY_CURRENT_CONFIG' = [Microsoft.Win32.RegistryHive]::CurrentConfig
    }

    $segments = $normalized -split '\\', 2
    $hiveName = $segments[0].TrimEnd(':').ToUpperInvariant()
    if (-not $hiveMap.ContainsKey($hiveName)) {
        return [PSCustomObject]@{ KeyExists = $false; ValueExists = $false; Value = $null; Error = "Unsupported registry hive '$($segments[0])'." }
    }
    $subPath = if ($segments.Count -gt 1) { $segments[1] } else { '' }
    $view = if ($Check32BitOn64System) { [Microsoft.Win32.RegistryView]::Registry32 } else { [Microsoft.Win32.RegistryView]::Registry64 }

    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hiveMap[$hiveName], $view)
        $key = if ($subPath) { $base.OpenSubKey($subPath) } else { $base }
        if (-not $key) {
            return [PSCustomObject]@{ KeyExists = $false; ValueExists = $false; Value = $null; Error = $null }
        }
        if ([string]::IsNullOrWhiteSpace($ValueName)) {
            return [PSCustomObject]@{ KeyExists = $true; ValueExists = $false; Value = $null; Error = $null }
        }
        $value = $key.GetValue($ValueName)
        return [PSCustomObject]@{
            KeyExists   = $true
            ValueExists = ($null -ne $value)
            Value       = $value
            Error       = $null
        }
    } catch {
        return [PSCustomObject]@{ KeyExists = $false; ValueExists = $false; Value = $null; Error = "$_" }
    }
}

function Test-RecipeDetection {
    <#
    .SYNOPSIS
    Evaluates a recipe's detection rule against the local machine.

    .DESCRIPTION
    Supports the file, msi, registry, and script detection types used by
    Yardstick recipes, matching Intune's evaluation semantics as closely as a
    local harness can.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable] $Recipe,
        [Parameter(Mandatory)][string] $Version,
        [AllowNull()][string] $ProductCode,
        [AllowNull()][string] $FileName,
        [Parameter(Mandatory)][string] $LogDirectory,
        [Parameter(Mandatory)][string] $LogName
    )

    function Get-RecipeValue([hashtable] $Table, [string] $Key) {
        if ($Table.ContainsKey($Key) -and $null -ne $Table[$Key] -and "$($Table[$Key])" -ne '') { return $Table[$Key] }
        return $null
    }

    $detectionType = [string](Get-RecipeValue $Recipe 'detectionType')
    $expectedVersion = [string](Get-RecipeValue $Recipe 'fileDetectionVersion')
    if ([string]::IsNullOrWhiteSpace($expectedVersion)) { $expectedVersion = $Version }
    $is32Bit = [bool](Get-RecipeValue $Recipe 'is32BitApp')

    # Yardstick's Set-ScriptPlaceholders substitutes <filename>, <version>, and
    # <productcode> in the install, uninstall, and detection scripts and in
    # registryDetectionKey only. The file detection fields are passed through
    # verbatim, so expanding tokens in them here would let a recipe pass the lab
    # and then fail in Intune.
    switch ($detectionType) {
        'file' {
            $path = Expand-RecipeString ([string](Get-RecipeValue $Recipe 'fileDetectionPath'))
            $name = Expand-RecipeString ([string](Get-RecipeValue $Recipe 'fileDetectionName'))
            $method = [string](Get-RecipeValue $Recipe 'fileDetectionMethod')
            $operator = [string](Get-RecipeValue $Recipe 'fileDetectionOperator')
            if (-not $operator) { $operator = 'greaterThanOrEqual' }
            $target = if ($name) { Join-Path $path $name } else { $path }
            $exists = Test-Path -LiteralPath $target

            if (-not $exists) {
                return [PSCustomObject]@{ Detected = $false; Rule = "file/$method"; Detail = "'$target' does not exist" }
            }

            switch ($method) {
                'exists' {
                    return [PSCustomObject]@{ Detected = $true; Rule = 'file/exists'; Detail = "'$target' exists" }
                }
                'version' {
                    $item = Get-Item -LiteralPath $target
                    $actual = [string]$item.VersionInfo.FileVersion
                    if ([string]::IsNullOrWhiteSpace($actual)) { $actual = [string]$item.VersionInfo.ProductVersion }
                    $ok = Compare-DetectionValue -Actual $actual -Expected $expectedVersion -Operator $operator -Kind version
                    return [PSCustomObject]@{ Detected = $ok; Rule = 'file/version'; Detail = "'$target' version '$actual' $operator '$expectedVersion'" }
                }
                'size' {
                    $sizeMb = [math]::Round((Get-Item -LiteralPath $target).Length / 1MB, 4)
                    $expected = [string](Get-RecipeValue $Recipe 'fileDetectionValue')
                    $ok = Compare-DetectionValue -Actual $sizeMb -Expected $expected -Operator $operator -Kind integer
                    return [PSCustomObject]@{ Detected = $ok; Rule = 'file/size'; Detail = "'$target' size ${sizeMb}MB $operator ${expected}MB" }
                }
                { $_ -in 'modified', 'created' } {
                    $item = Get-Item -LiteralPath $target
                    $actual = if ($method -eq 'modified') { $item.LastWriteTime } else { $item.CreationTime }
                    $expected = [string](Get-RecipeValue $Recipe 'fileDetectionDateTime')
                    $ok = Compare-DetectionValue -Actual $actual -Expected $expected -Operator $operator -Kind datetime
                    return [PSCustomObject]@{ Detected = $ok; Rule = "file/$method"; Detail = "'$target' $method '$actual' $operator '$expected'" }
                }
                default {
                    return [PSCustomObject]@{ Detected = $exists; Rule = "file/$method"; Detail = "Unknown fileDetectionMethod '$method'; fell back to existence of '$target'" }
                }
            }
        }
        'msi' {
            if ([string]::IsNullOrWhiteSpace($ProductCode)) {
                return [PSCustomObject]@{ Detected = $false; Rule = 'msi'; Detail = 'No MSI product code was resolved from the installer' }
            }
            $match = @(Get-InstalledPackage | Where-Object { $_.SubKeyName -ieq $ProductCode })
            if ($match.Count -eq 0) {
                return [PSCustomObject]@{ Detected = $false; Rule = 'msi'; Detail = "Product code '$ProductCode' is not registered" }
            }
            $installed = [string]$match[0].DisplayVersion
            $ok = Compare-DetectionValue -Actual $installed -Expected $expectedVersion -Operator 'greaterThanOrEqual' -Kind version
            return [PSCustomObject]@{ Detected = $ok; Rule = 'msi'; Detail = "Product code '$ProductCode' version '$installed' >= '$expectedVersion'" }
        }
        'registry' {
            $keyPath = Expand-RecipeString ([string](Expand-RecipeToken -Value ([string](Get-RecipeValue $Recipe 'registryDetectionKey')) -FileName $FileName -Version $Version -ProductCode $ProductCode))
            $valueName = [string](Get-RecipeValue $Recipe 'registryDetectionValueName')
            $method = [string](Get-RecipeValue $Recipe 'registryDetectionMethod')
            $operator = [string](Get-RecipeValue $Recipe 'registryDetectionOperator')
            if (-not $operator) { $operator = 'greaterThanOrEqual' }
            $expected = [string](Get-RecipeValue $Recipe 'registryDetectionValue')
            if ([string]::IsNullOrWhiteSpace($expected) -and $method -eq 'version') { $expected = $expectedVersion }

            $lookup = Get-RegistryValueForDetection -KeyPath $keyPath -ValueName $valueName -Check32BitOn64System:$is32Bit
            if ($lookup.Error) {
                return [PSCustomObject]@{ Detected = $false; Rule = "registry/$method"; Detail = $lookup.Error }
            }
            switch ($method) {
                'exists' {
                    $ok = if ($valueName) { $lookup.ValueExists } else { $lookup.KeyExists }
                    return [PSCustomObject]@{ Detected = $ok; Rule = 'registry/exists'; Detail = "'$keyPath' value '$valueName' exists=$ok" }
                }
                'version' {
                    $ok = Compare-DetectionValue -Actual $lookup.Value -Expected $expected -Operator $operator -Kind version
                    return [PSCustomObject]@{ Detected = $ok; Rule = 'registry/version'; Detail = "'$keyPath\$valueName' = '$($lookup.Value)' $operator '$expected'" }
                }
                'integer' {
                    $ok = Compare-DetectionValue -Actual $lookup.Value -Expected $expected -Operator $operator -Kind integer
                    return [PSCustomObject]@{ Detected = $ok; Rule = 'registry/integer'; Detail = "'$keyPath\$valueName' = '$($lookup.Value)' $operator '$expected'" }
                }
                'string' {
                    $ok = Compare-DetectionValue -Actual $lookup.Value -Expected $expected -Operator $operator -Kind string
                    return [PSCustomObject]@{ Detected = $ok; Rule = 'registry/string'; Detail = "'$keyPath\$valueName' = '$($lookup.Value)' $operator '$expected'" }
                }
                default {
                    return [PSCustomObject]@{ Detected = $lookup.KeyExists; Rule = "registry/$method"; Detail = "Unknown registryDetectionMethod '$method'; fell back to key existence" }
                }
            }
        }
        'script' {
            $content = [string](Get-RecipeValue $Recipe 'detectionScript')
            if ([string]::IsNullOrWhiteSpace($content)) {
                return [PSCustomObject]@{ Detected = $false; Rule = 'script'; Detail = 'detectionType is script but detectionScript is empty' }
            }
            $content = Expand-RecipeToken -Value $content -FileName $FileName -Version $Version -ProductCode $ProductCode
            $runAs32 = [bool](Get-RecipeValue $Recipe 'detectionScriptRunAs32Bit')
            $result = Invoke-DetectionScript -ScriptContent $content -LogDirectory $LogDirectory -LogName $LogName -RunAs32Bit:$runAs32
            return [PSCustomObject]@{ Detected = $result.Detected; Rule = 'script'; Detail = $result.Detail }
        }
        default {
            return [PSCustomObject]@{ Detected = $false; Rule = "unknown/$detectionType"; Detail = "Unsupported detectionType '$detectionType'" }
        }
    }
}

Export-ModuleMember -Function Test-IsElevated, Expand-RecipeString, Expand-RecipeToken,
ConvertTo-ComparableVersion, Compare-DetectionValue, Get-InstalledPackage, Get-SystemSnapshot,
Compare-SystemSnapshot, Invoke-TestProcess, Invoke-CommandLine, Invoke-DetectionScript,
Get-RegistryValueForDetection, Test-RecipeDetection
