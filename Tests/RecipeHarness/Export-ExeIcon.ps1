<#
.SYNOPSIS
Extracts the largest embedded application icon from a Windows executable and
writes it as a PNG suitable for a Yardstick recipe icon.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $ExePath,
    [Parameter(Mandatory)][string] $OutputPath,
    [int] $PreferredSize = 256
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Add-Type -AssemblyName System.Drawing

$signature = @'
[DllImport("user32.dll", CharSet = CharSet.Unicode)]
public static extern int PrivateExtractIcons(string lpszFile, int nIconIndex, int cxIcon, int cyIcon, IntPtr[] phicon, int[] piconid, int nIcons, int flags);
[DllImport("user32.dll")]
public static extern bool DestroyIcon(IntPtr hIcon);
'@
if (-not ('Win32IconExtractor' -as [type])) {
    Add-Type -MemberDefinition $signature -Name 'Win32IconExtractor' -Namespace '' -PassThru | Out-Null
}

$ExePath = (Resolve-Path -LiteralPath $ExePath).Path

foreach ($size in @($PreferredSize, 128, 64, 48, 32)) {
    $handles = New-Object IntPtr[] 1
    $ids = New-Object int[] 1
    $count = [Win32IconExtractor]::PrivateExtractIcons($ExePath, 0, $size, $size, $handles, $ids, 1, 0)
    if ($count -le 0 -or $handles[0] -eq [IntPtr]::Zero) { continue }

    try {
        $icon = [System.Drawing.Icon]::FromHandle($handles[0])
        $bitmap = $icon.ToBitmap()
        try {
            $bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        } finally {
            $bitmap.Dispose()
            $icon.Dispose()
        }
    } finally {
        [Win32IconExtractor]::DestroyIcon($handles[0]) | Out-Null
    }

    return [PSCustomObject]@{
        Source = $ExePath
        Output = (Resolve-Path -LiteralPath $OutputPath).Path
        Size   = $size
    }
}

throw "No embedded icon could be extracted from '$ExePath'."
