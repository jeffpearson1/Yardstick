<#
.SYNOPSIS
    Read-only diagnostic for the Yardstick mGear deployment.

.DESCRIPTION
    Gathers evidence about where the mGear module descriptors actually landed,
    how CommonProgramFiles resolves in each process bitness, what Maya reports
    as its module search path, and what the Intune Management Extension logged.

    This script makes NO changes. It only reads files, environment variables and
    logs, and writes a single report to $env:TEMP.

    Run it in an ADMINISTRATOR PowerShell so it can read Program Files and the
    IntuneManagementExtension logs.

.PARAMETER ExpectedVersion
    The mGear package version to check for. Defaults to 5.3.5.

.EXAMPLE
    .\Test-MgearDeployment.ps1
    .\Test-MgearDeployment.ps1 -ExpectedVersion 5.3.5
#>
[CmdletBinding()]
param(
    [string]$ExpectedVersion = '5.3.5'
)

$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'

# Resolve well-known folders through .NET rather than environment variables.
# Some shells (git-bash, WSL interop, restricted service contexts) drop the
# Program Files / ProgramData variables, which would silently skew this report.
$TempDir      = [IO.Path]::GetTempPath()
$ProgramData  = [Environment]::GetFolderPath('CommonApplicationData')
$CommonNative = [Environment]::GetFolderPath('CommonProgramFiles')
$CommonX86    = [Environment]::GetFolderPath('CommonProgramFilesX86')
$WinDir       = [Environment]::GetFolderPath('Windows')
if (-not $ProgramData)  { $ProgramData  = 'C:\ProgramData' }
if (-not $CommonNative) { $CommonNative = 'C:\Program Files\Common Files' }
if (-not $CommonX86)    { $CommonX86    = 'C:\Program Files (x86)\Common Files' }
if (-not $WinDir)       { $WinDir       = 'C:\Windows' }

$reportPath = Join-Path $TempDir ("mgear-diagnostic-{0}-{1}.txt" -f [Environment]::MachineName, (Get-Date -Format 'yyyyMMdd-HHmmss'))
$lines = [System.Collections.Generic.List[string]]::new()

function Add-Line { param([string]$Text = '') ; $lines.Add($Text) ; Write-Host $Text }
function Add-Section {
    param([string]$Title)
    Add-Line ''
    Add-Line ('=' * 78)
    Add-Line "  $Title"
    Add-Line ('=' * 78)
}
function Add-Kv { param([string]$Key, $Value) ; Add-Line ("  {0,-34} {1}" -f $Key, $Value) }

Add-Line "mGear deployment diagnostic"
Add-Line "Computer      : $([Environment]::MachineName)"
Add-Line "User          : $([Environment]::UserName)"
Add-Line "Timestamp     : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')"
Add-Line "Expecting     : mGear $ExpectedVersion"
Add-Line "OS            : $((Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption)"
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
Add-Line "Elevated      : $isAdmin"
if (-not $isAdmin) {
    Add-Line "  ** WARNING: not elevated. Program Files and IME log checks may be incomplete. **"
}


# ---------------------------------------------------------------------------
Add-Section '1. Process bitness and Program Files resolution'
# ---------------------------------------------------------------------------
# The core question: does CommonProgramFiles resolve differently in a 32-bit
# process on this machine? If it does, a 32-bit installer writes to the wrong tree.

Add-Line 'This host:'
Add-Kv 'Is64BitProcess'          ([Environment]::Is64BitProcess)
Add-Kv 'Is64BitOperatingSystem'  ([Environment]::Is64BitOperatingSystem)
Add-Kv 'env:CommonProgramFiles'  ($(if ($env:CommonProgramFiles) { $env:CommonProgramFiles } else { '(not set)' }))
Add-Kv 'env:CommonProgramW6432'  ($(if ($env:CommonProgramW6432) { $env:CommonProgramW6432 } else { '(not set)' }))
Add-Kv 'Resolved native Common'  $CommonNative
Add-Kv 'Resolved x86 Common'     $CommonX86

Add-Line ''
Add-Line 'Spawned child processes (this is the mechanism under test):'

$probeResults = @{}

function Get-ChildEnvProbe {
    param([string]$CmdPath, [string]$Label, [string]$Key)
    if (-not (Test-Path -LiteralPath $CmdPath)) {
        Add-Line "  $Label"
        Add-Kv '    (result)' 'cmd.exe not present on this machine'
        return
    }
    $out = Join-Path $TempDir ("mgear-envprobe-{0}.txt" -f ([guid]::NewGuid().ToString('N')))
    try {
        # 'set' is a cmd builtin; /c set dumps the child's own environment block.
        $null = Start-Process -FilePath $CmdPath -ArgumentList '/c', "set > `"$out`"" `
                              -Wait -WindowStyle Hidden -PassThru -ErrorAction Stop
        $childEnv = @{}
        Get-Content -LiteralPath $out -ErrorAction Stop | ForEach-Object {
            if ($_ -match '^([^=]+)=(.*)$') { $childEnv[$matches[1]] = $matches[2] }
        }
        Add-Line "  $Label"
        foreach ($v in 'CommonProgramFiles', 'CommonProgramW6432', 'ProgramFiles', 'ProgramW6432') {
            $val = if ($childEnv.ContainsKey($v)) { $childEnv[$v] } else { '(not set)' }
            Add-Kv "    $v" $val
        }
        $probeResults[$Key] = $childEnv
    }
    catch {
        Add-Line "  $Label"
        Add-Kv '    (result)' "probe failed: $($_.Exception.Message)"
    }
    finally { Remove-Item -LiteralPath $out -Force -ErrorAction SilentlyContinue }
}

Get-ChildEnvProbe -CmdPath "$WinDir\SysWOW64\cmd.exe" -Label '32-bit child (SysWOW64\cmd.exe):' -Key '32'
Get-ChildEnvProbe -CmdPath "$WinDir\System32\cmd.exe" -Label '64-bit child (System32\cmd.exe):' -Key '64'

$c32 = $probeResults['32']; $c64 = $probeResults['64']
if ($c32 -and $c64) {
    if (-not $c32.ContainsKey('CommonProgramFiles') -and -not $c64.ContainsKey('CommonProgramFiles')) {
        Add-Line ''
        Add-Line '  ** These variables are unset in BOTH children, which means the shell you'
        Add-Line '     launched this script from stripped them (git-bash / WSL interop do this).'
        Add-Line '     Re-run from a plain elevated PowerShell prompt for a valid section 1. **'
    }
    elseif ($c32['CommonProgramFiles'] -ne $c64['CommonProgramFiles']) {
        Add-Line ''
        Add-Line '  >> CONFIRMED: a 32-bit child resolves CommonProgramFiles differently.'
        Add-Line "     32-bit sees: $($c32['CommonProgramFiles'])"
        Add-Line "     64-bit sees: $($c64['CommonProgramFiles'])"
        Add-Line '     A 32-bit install script therefore writes to the wrong tree.'
    }
    else {
        Add-Line ''
        Add-Line '  >> Both bitnesses resolve CommonProgramFiles identically on this box.'
        Add-Line '     That would rule OUT the bitness explanation.'
    }
}


# ---------------------------------------------------------------------------
Add-Section '2. Module descriptor trees - where did mGear.mod actually land?'
# ---------------------------------------------------------------------------

$trees = [ordered]@{
    'NATIVE (what Maya reads)' = Join-Path $CommonNative 'Autodesk Shared\Modules\maya'
    'x86 (the wrong tree)'     = Join-Path $CommonX86    'Autodesk Shared\Modules\maya'
}

$foundNative = @()
$foundX86    = @()

foreach ($label in $trees.Keys) {
    $root = $trees[$label]
    Add-Line ''
    Add-Line "$label"
    Add-Line "  $root"
    if (-not (Test-Path -LiteralPath $root)) {
        Add-Line '    -> folder does not exist'
        continue
    }

    # Show the immediate children so we can see whether year subfolders exist.
    $kids = @(Get-ChildItem -LiteralPath $root -Force -ErrorAction SilentlyContinue)
    if ($kids.Count -eq 0) { Add-Line '    -> folder exists but is EMPTY' }
    else { Add-Line ("    children: " + (($kids | ForEach-Object { $_.Name }) -join ', ')) }

    $mods = @(Get-ChildItem -LiteralPath $root -Filter 'mGear.mod' -File -Recurse -ErrorAction SilentlyContinue)
    if ($mods.Count -eq 0) {
        Add-Line '    -> no mGear.mod found anywhere under this root'
    }
    foreach ($m in $mods) {
        $head = ''
        try { $head = (Get-Content -LiteralPath $m.FullName -TotalCount 1 -ErrorAction Stop) } catch { $head = "(unreadable: $($_.Exception.Message))" }
        Add-Line "    FOUND: $($m.FullName)"
        Add-Line "           modified : $($m.LastWriteTime)"
        Add-Line "           header   : $head"
        $managed = $head -match '^# Managed by Yardstick; package version (.+?)\s*$'
        Add-Line "           yardstick-managed : $managed$(if ($managed) { " (version $($matches[1]))" })"
        if ($label -like 'NATIVE*') { $foundNative += $m } else { $foundX86 += $m }
    }
}


# ---------------------------------------------------------------------------
Add-Section '3. Payload root - C:\ProgramData\mGear'
# ---------------------------------------------------------------------------

$installRoot = Join-Path $ProgramData 'mGear'
Add-Kv 'Path' $installRoot
Add-Kv 'Exists' (Test-Path -LiteralPath $installRoot)

$markerVersion = $null
if (Test-Path -LiteralPath $installRoot) {
    $kids = @(Get-ChildItem -LiteralPath $installRoot -Force -ErrorAction SilentlyContinue)
    Add-Kv 'Children' (($kids | ForEach-Object { $_.Name }) -join ', ')

    $marker = Join-Path $installRoot '.yardstick-version'
    if (Test-Path -LiteralPath $marker) {
        $markerVersion = (Get-Content -LiteralPath $marker -Raw -ErrorAction SilentlyContinue).Trim()
        Add-Kv '.yardstick-version' $markerVersion
    } else {
        Add-Kv '.yardstick-version' '(MISSING)'
    }

    $manifest = Join-Path $installRoot 'mGear.mod'
    Add-Kv 'mGear.mod present' (Test-Path -LiteralPath $manifest)
    if (Test-Path -LiteralPath $manifest) {
        Add-Line ''
        Add-Line '  First 6 lines of the installed mGear.mod:'
        Get-Content -LiteralPath $manifest -TotalCount 6 -ErrorAction SilentlyContinue |
            ForEach-Object { Add-Line "    | $_" }
    }

    $platformRoot = Join-Path $installRoot 'platforms'
    if (Test-Path -LiteralPath $platformRoot) {
        $years = @(Get-ChildItem -LiteralPath $platformRoot -Directory -ErrorAction SilentlyContinue |
                   Where-Object { $_.Name -match '^\d{4}$' -and (Test-Path -LiteralPath (Join-Path $_.FullName 'windows\x64\plug-ins')) })
        Add-Line ''
        Add-Kv 'Windows x64 platform years' (($years | ForEach-Object { $_.Name }) -join ', ')
    }
}


# ---------------------------------------------------------------------------
Add-Section '4. Detection script result (as Intune would evaluate it)'
# ---------------------------------------------------------------------------
# Replicates Recipes/mgear.yaml detectionScript verbatim, in-process.

$detected = $false
$expected = $null
if (-not [version]::TryParse($ExpectedVersion, [ref]$expected)) {
    Add-Line "  Expected version '$ExpectedVersion' is not parseable -> would report Not Detected"
}
else {
    $markerPath   = Join-Path $installRoot '.yardstick-version'
    $manifestPath = Join-Path $installRoot 'mGear.mod'
    if (-not (Test-Path -LiteralPath $markerPath) -or -not (Test-Path -LiteralPath $manifestPath)) {
        Add-Line '  Marker or manifest missing under C:\ProgramData\mGear -> Not Detected'
    }
    else {
        $installed    = $null
        $installedText = (Get-Content -LiteralPath $markerPath -Raw).Trim()
        if (-not [version]::TryParse($installedText, [ref]$installed) -or $installed -lt $expected) {
            Add-Line "  Installed marker '$installedText' unparseable or older than $ExpectedVersion -> Not Detected"
        }
        else {
            # Mirrors the recipe's own expression. If the shell stripped the
            # variables, fall back to the resolved path and say so.
            $commonFiles = if ($env:CommonProgramW6432) { $env:CommonProgramW6432 } else { $env:CommonProgramFiles }
            if (-not $commonFiles) {
                $commonFiles = $CommonNative
                Add-Line '  (env vars unavailable in this shell; using resolved native path)'
            }
            $moduleRoot  = Join-Path $commonFiles 'Autodesk Shared\Modules\maya'
            Add-Kv 'Detection would search' $moduleRoot
            $descriptor = Get-ChildItem -LiteralPath $moduleRoot -Filter 'mGear.mod' -File -Recurse -ErrorAction SilentlyContinue |
                Where-Object {
                    (Get-Content -LiteralPath $_.FullName -Raw) -match
                        "(?m)^# Managed by Yardstick; package version $([regex]::Escape($installedText))\r?$"
                } | Select-Object -First 1
            if ($descriptor) {
                $detected = $true
                Add-Line "  RESULT: Detected  (matched $($descriptor.FullName))"
            } else {
                Add-Line '  RESULT: Not Detected  (no matching descriptor in the native tree)'
            }
        }
    }
}


# ---------------------------------------------------------------------------
Add-Section '5. Maya installations and their real module search path'
# ---------------------------------------------------------------------------
# Settles whether Maya scans ...\Modules\maya\<year> or the unversioned folder.

$autodeskRoot = Join-Path ([Environment]::GetFolderPath('ProgramFiles')) 'Autodesk'
if (-not (Test-Path -LiteralPath $autodeskRoot)) { $autodeskRoot = 'C:\Program Files\Autodesk' }
Add-Kv 'Searching' $autodeskRoot
$mayaRoots = @(Get-ChildItem -LiteralPath $autodeskRoot -Directory -Filter 'Maya*' -ErrorAction SilentlyContinue)
if ($mayaRoots.Count -eq 0) {
    Add-Line "  No Maya installations found under $autodeskRoot"
    Add-Line '  (If Maya is installed elsewhere, note the path when you report back.)'
}
foreach ($maya in $mayaRoots) {
    Add-Line ''
    Add-Line "  $($maya.Name)  ->  $($maya.FullName)"
    $mayapy = Join-Path $maya.FullName 'bin\mayapy.exe'
    if (-not (Test-Path -LiteralPath $mayapy)) {
        Add-Line '    mayapy.exe not found; skipping module path query'
        continue
    }

    # Ask Maya itself. Standalone init is slow, so cap it.
    $py  = Join-Path $TempDir ("mgear-modpath-{0}.py"  -f ([guid]::NewGuid().ToString('N')))
    $out = Join-Path $TempDir ("mgear-modpath-{0}.txt" -f ([guid]::NewGuid().ToString('N')))
    @'
import os, sys
try:
    import maya.standalone
    maya.standalone.initialize(name="python")
except Exception as e:
    sys.stderr.write("standalone init failed: %s\n" % e)
paths = os.environ.get("MAYA_MODULE_PATH", "")
for p in paths.split(os.pathsep):
    if p.strip():
        print("MODPATH\t%s\t%s" % ("EXISTS" if os.path.isdir(p) else "missing", p))
'@ | Set-Content -LiteralPath $py -Encoding ASCII

    try {
        $p = Start-Process -FilePath $mayapy -ArgumentList "`"$py`"" -Wait -WindowStyle Hidden `
                           -RedirectStandardOutput $out -RedirectStandardError "$out.err" -PassThru -ErrorAction Stop
        $modLines = @(Get-Content -LiteralPath $out -ErrorAction SilentlyContinue | Where-Object { $_ -like 'MODPATH*' })
        if ($modLines.Count -eq 0) {
            Add-Line '    (Maya returned no MAYA_MODULE_PATH; see stderr below)'
            Get-Content -LiteralPath "$out.err" -TotalCount 5 -ErrorAction SilentlyContinue |
                ForEach-Object { Add-Line "      ! $_" }
        }
        foreach ($l in $modLines) {
            $parts = $l -split "`t"
            Add-Line ("    [{0,-7}] {1}" -f $parts[1], $parts[2])
        }
    }
    catch { Add-Line "    mayapy query failed: $($_.Exception.Message)" }
    finally { Remove-Item -LiteralPath $py, $out, "$out.err" -Force -ErrorAction SilentlyContinue }
}


# ---------------------------------------------------------------------------
Add-Section '6. Intune Management Extension evidence'
# ---------------------------------------------------------------------------

$imeExe = Join-Path $CommonX86.Replace('\Common Files', '') 'Microsoft Intune Management Extension\Microsoft.Management.Services.IntuneWindowsAgent.exe'
if (-not (Test-Path -LiteralPath $imeExe)) {
    $imeExe = 'C:\Program Files (x86)\Microsoft Intune Management Extension\Microsoft.Management.Services.IntuneWindowsAgent.exe'
}
Add-Kv 'IME executable present' (Test-Path -LiteralPath $imeExe)
if (Test-Path -LiteralPath $imeExe) {
    # A PE32 (not PE32+) header confirms the agent - and therefore its child
    # install processes - is 32-bit.
    try {
        $fs = [IO.File]::OpenRead($imeExe); $br = [IO.BinaryReader]::new($fs)
        $fs.Position = 0x3C; $peOffset = $br.ReadInt32(); $fs.Position = $peOffset + 4
        $machine = $br.ReadUInt16(); $br.Close(); $fs.Close()
        $arch = switch ($machine) { 0x014c { 'x86 (32-bit)' } 0x8664 { 'x64 (64-bit)' } 0xAA64 { 'ARM64' } default { "unknown (0x{0:X})" -f $machine } }
        Add-Kv 'IME architecture' $arch
    } catch { Add-Kv 'IME architecture' "(could not read: $($_.Exception.Message))" }
}

$logDir = Join-Path $ProgramData 'Microsoft\IntuneManagementExtension\Logs'
Add-Kv 'Log directory' $logDir
if (Test-Path -LiteralPath $logDir) {
    $hits = @(Select-String -Path (Join-Path $logDir 'IntuneManagementExtension*.log') `
                            -Pattern 'mgear' -SimpleMatch -ErrorAction SilentlyContinue |
              Select-Object -Last 40)
    if ($hits.Count -eq 0) {
        Add-Line '  No lines mentioning "mgear" in the IME logs.'
    } else {
        Add-Line "  Last $($hits.Count) IME log lines mentioning mgear:"
        foreach ($h in $hits) { Add-Line ("    [{0}:{1}] {2}" -f $h.Filename, $h.LineNumber, $h.Line.Trim()) }
    }
} else {
    Add-Line '  IME log directory not present.'
}


# ---------------------------------------------------------------------------
Add-Section '7. Verdict'
# ---------------------------------------------------------------------------

Add-Kv 'Descriptors in NATIVE tree' $foundNative.Count
Add-Kv 'Descriptors in x86 tree'    $foundX86.Count
Add-Kv 'Payload marker version'     ($(if ($markerVersion) { $markerVersion } else { '(none)' }))
Add-Kv 'Detection result'           ($(if ($detected) { 'Detected' } else { 'Not Detected' }))
Add-Line ''

if ($foundX86.Count -gt 0 -and $foundNative.Count -eq 0) {
    Add-Line '  >> CONFIRMS the 32-bit diagnosis: descriptors exist ONLY in the'
    Add-Line '     "Program Files (x86)" tree, where Maya never looks.'
    Add-Line '     The x86 copies are orphaned - the fixed uninstaller will not'
    Add-Line '     remove them, so clear that folder by hand.'
}
elseif ($foundNative.Count -gt 0 -and -not $detected) {
    Add-Line '  >> NOT the bitness issue. Descriptors are in the right tree but'
    Add-Line '     detection still fails - most likely a version-marker or header'
    Add-Line '     mismatch. Compare section 2 headers against section 3 marker.'
}
elseif ($foundNative.Count -eq 0 -and $foundX86.Count -eq 0 -and -not (Test-Path -LiteralPath $installRoot)) {
    Add-Line '  >> Nothing installed at all. The install script never ran to'
    Add-Line '     completion - check section 6 for the IME exit code.'
}
elseif ($foundNative.Count -eq 0 -and $foundX86.Count -eq 0) {
    Add-Line '  >> Payload exists under C:\ProgramData\mGear but NO descriptors were'
    Add-Line '     written to either tree. The install failed partway - check'
    Add-Line '     section 6 for the error.'
}
elseif ($detected) {
    Add-Line '  >> Healthy: descriptors are in the native tree and detection passes.'
}

Add-Line ''
Add-Line ('-' * 78)
Set-Content -LiteralPath $reportPath -Value $lines -Encoding UTF8
Write-Host ''
Write-Host "Report written to: $reportPath" -ForegroundColor Cyan
Write-Host 'Send that file back (it contains no credentials).' -ForegroundColor Cyan
