#Requires -Version 7.0
<#
.SYNOPSIS
Makes the setup changes that need administrator rights. Used by Install-Yardstick.ps1.

.DESCRIPTION
Install-Yardstick.ps1 runs as the user who runs Yardstick, without administrator
rights, and collects the few changes that need them. Invoke-YardstickElevatedSetup
then starts this script once with "Run as administrator" to make them all. It
writes nothing else: no log, and nothing in the Yardstick folder beyond what the
actions themselves change.

.PARAMETER RequestPath
JSON list of actions (see Invoke-YardstickElevatedAction).

.PARAMETER ResultPath
Where to write the result of each action.

.PARAMETER Key
Base64 key that decrypts the passwords in the request.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$RequestPath,
    [Parameter(Mandatory)][string]$ResultPath,
    [string]$Key
)

$ErrorActionPreference = 'Stop'
$results = [System.Collections.Generic.List[object]]::new()
$fatal = $null
try {
    Import-Module (Join-Path $PSScriptRoot '..\Yardstick.psd1')
    $keyBytes = if ($Key) { [Convert]::FromBase64String($Key) } else { $null }
    foreach ($action in @(Get-Content -LiteralPath $RequestPath -Raw | ConvertFrom-Json)) {
        try {
            Invoke-YardstickElevatedAction -Action $action -Key $keyBytes
            $results.Add([pscustomobject]@{ Success = $true; Message = $null })
        } catch {
            $results.Add([pscustomobject]@{ Success = $false; Message = $_.Exception.Message })
        }
    }
} catch {
    $fatal = $_.Exception.Message
}
ConvertTo-Json -InputObject ([pscustomobject]@{ Results = @($results); Error = $fatal }) -Depth 4 |
    Set-Content -LiteralPath $ResultPath -Encoding utf8
