<#
.SYNOPSIS
Secure storage and lifecycle management for the Intune app-registration
credentials (TenantID / ClientID / ClientSecret) used by Yardstick.

.DESCRIPTION
Credentials live in Windows Credential Manager as a single generic credential
rather than in plaintext inside preferences.yaml. The credential is stored under
a configurable target name (preferences key `credentialTarget`, default
"Yardstick:IntuneGraph") with the ClientID as the credential user name and a
small JSON document as the credential blob:

    { "tenantId": "...", "clientSecret": "...", "secretExpiresOn": "...", "lastNotifiedOn": "..." }

Generic credentials are encrypted by Windows under the storing user's profile,
so only that user (on that machine) can read them back.

The client secret's expiration is discovered from Microsoft Graph when the app
registration has Application.Read.All, matching the secret to its
passwordCredential by hint. When Graph cannot be queried, the expiration entered
at setup time is used instead.
#>

$Script:DefaultCredentialTarget = 'Yardstick:IntuneGraph'
$Script:DefaultExpirationWarningDays = 30
$Script:DefaultNotificationIntervalHours = 24

# ERROR_NOT_FOUND from CredReadW when no credential exists for the target.
$Script:CredNotFound = 1168


function Write-CredentialLog {
    <#
    .SYNOPSIS
    Logs through Yardstick's Write-Log when available, otherwise to the console.
    #>
    param(
        [string]$Content,
        [switch]$Warning
    )

    if ($Warning) { Write-Warning $Content }

    if (Get-Command -Name Write-Log -ErrorAction SilentlyContinue) {
        Write-Log ($(if ($Warning) { "WARNING: $Content" } else { $Content }))
    } elseif (-not $Warning) {
        Write-Host "$(Get-Date -Format 'MM/dd/yyyy HH:mm:ss') - $Content"
    }
}


function Initialize-YardstickCredentialType {
    <#
    .SYNOPSIS
    Compiles the advapi32 Credential Manager P/Invoke wrapper once per session.

    .DESCRIPTION
    Windows Credential Manager has no built-in PowerShell cmdlets, and the
    third-party modules that wrap it are optional in this repo. Compiling the
    interop here keeps credential storage dependency-free.
    #>
    if ('YardstickCredentialNative' -as [type]) { return }

    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class YardstickCredentialNative
{
    private const uint CRED_TYPE_GENERIC = 1;
    private const uint CRED_PERSIST_LOCAL_MACHINE = 2;
    public const int MAX_CREDENTIAL_BLOB_SIZE = 2560;
    public const int ERROR_NOT_FOUND = 1168;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct CREDENTIAL
    {
        public uint Flags;
        public uint Type;
        public IntPtr TargetName;
        public IntPtr Comment;
        public long LastWritten;
        public uint CredentialBlobSize;
        public IntPtr CredentialBlob;
        public uint Persist;
        public uint AttributeCount;
        public IntPtr Attributes;
        public IntPtr TargetAlias;
        public IntPtr UserName;
    }

    [DllImport("advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CredRead(string target, uint type, uint reservedFlag, out IntPtr credentialPtr);

    [DllImport("advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CredWrite([In] ref CREDENTIAL userCredential, uint flags);

    [DllImport("advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CredDelete(string target, uint type, uint flags);

    [DllImport("advapi32.dll", EntryPoint = "CredFree")]
    private static extern void CredFree(IntPtr cred);

    public static bool Read(string target, out string userName, out string secret)
    {
        userName = null;
        secret = null;

        IntPtr credPtr = IntPtr.Zero;
        if (!CredRead(target, CRED_TYPE_GENERIC, 0, out credPtr))
        {
            int err = Marshal.GetLastWin32Error();
            if (err == ERROR_NOT_FOUND) { return false; }
            throw new InvalidOperationException("CredReadW failed for target '" + target + "' with Win32 error " + err + ".");
        }

        try
        {
            CREDENTIAL cred = (CREDENTIAL)Marshal.PtrToStructure(credPtr, typeof(CREDENTIAL));
            if (cred.UserName != IntPtr.Zero)
            {
                userName = Marshal.PtrToStringUni(cred.UserName);
            }
            if (cred.CredentialBlob != IntPtr.Zero && cred.CredentialBlobSize > 1)
            {
                secret = Marshal.PtrToStringUni(cred.CredentialBlob, (int)(cred.CredentialBlobSize / 2));
            }
            return true;
        }
        finally
        {
            CredFree(credPtr);
        }
    }

    public static void Write(string target, string userName, string secret, string comment)
    {
        byte[] blob = Encoding.Unicode.GetBytes(secret == null ? string.Empty : secret);
        if (blob.Length > MAX_CREDENTIAL_BLOB_SIZE)
        {
            throw new ArgumentException("Credential blob is " + blob.Length + " bytes, which exceeds the Windows limit of " + MAX_CREDENTIAL_BLOB_SIZE + " bytes.");
        }

        IntPtr blobPtr = Marshal.AllocHGlobal(blob.Length == 0 ? 1 : blob.Length);
        IntPtr targetPtr = Marshal.StringToHGlobalUni(target);
        IntPtr userPtr = Marshal.StringToHGlobalUni(string.IsNullOrEmpty(userName) ? " " : userName);
        IntPtr commentPtr = string.IsNullOrEmpty(comment) ? IntPtr.Zero : Marshal.StringToHGlobalUni(comment);

        try
        {
            Marshal.Copy(blob, 0, blobPtr, blob.Length);

            CREDENTIAL cred = new CREDENTIAL();
            cred.Flags = 0;
            cred.Type = CRED_TYPE_GENERIC;
            cred.TargetName = targetPtr;
            cred.Comment = commentPtr;
            cred.CredentialBlobSize = (uint)blob.Length;
            cred.CredentialBlob = blobPtr;
            cred.Persist = CRED_PERSIST_LOCAL_MACHINE;
            cred.AttributeCount = 0;
            cred.Attributes = IntPtr.Zero;
            cred.TargetAlias = IntPtr.Zero;
            cred.UserName = userPtr;

            if (!CredWrite(ref cred, 0))
            {
                throw new InvalidOperationException("CredWriteW failed for target '" + target + "' with Win32 error " + Marshal.GetLastWin32Error() + ".");
            }
        }
        finally
        {
            Array.Clear(blob, 0, blob.Length);
            for (int i = 0; i < blob.Length; i++) { Marshal.WriteByte(blobPtr, i, 0); }
            Marshal.FreeHGlobal(blobPtr);
            Marshal.FreeHGlobal(targetPtr);
            Marshal.FreeHGlobal(userPtr);
            if (commentPtr != IntPtr.Zero) { Marshal.FreeHGlobal(commentPtr); }
        }
    }

    public static bool Delete(string target)
    {
        if (CredDelete(target, CRED_TYPE_GENERIC, 0)) { return true; }
        int err = Marshal.GetLastWin32Error();
        if (err == ERROR_NOT_FOUND) { return false; }
        throw new InvalidOperationException("CredDeleteW failed for target '" + target + "' with Win32 error " + err + ".");
    }
}
'@
}


function Get-YardstickCredentialTarget {
    <#
    .SYNOPSIS
    Resolves the Credential Manager target name from preferences.

    .PARAMETER Preferences
    Parsed preferences.yaml hashtable. Optional.
    #>
    param(
        $Preferences
    )

    if ($Preferences -and -not [string]::IsNullOrWhiteSpace([string]$Preferences.credentialTarget)) {
        return [string]$Preferences.credentialTarget
    }
    return $Script:DefaultCredentialTarget
}


function ConvertTo-PlainSecret {
    <#
    .SYNOPSIS
    Normalizes a SecureString or String secret to a plain String.
    #>
    param($Secret)

    if ($null -eq $Secret) { return $null }
    if ($Secret -is [System.Security.SecureString]) {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secret)
        try {
            return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        } finally {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
    return [string]$Secret
}


function ConvertTo-CredentialDateTime {
    <#
    .SYNOPSIS
    Parses a stored or Graph-supplied timestamp into a local DateTime, or $null.
    #>
    param($Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }

    # ConvertFrom-Json hydrates ISO 8601 strings into DateTime already, in UTC.
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [System.DateTimeKind]::Utc) { return $Value.ToLocalTime() }
        return $Value
    }

    [datetime]$parsed = [datetime]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::RoundtripKind
    if ([datetime]::TryParse([string]$Value, [cultureinfo]::InvariantCulture, $styles, [ref]$parsed)) {
        # Stored and Graph timestamps carry a Z; a bare date typed by an operator does not.
        if ($parsed.Kind -eq [System.DateTimeKind]::Utc) { return $parsed.ToLocalTime() }
        return $parsed
    }
    return $null
}


function Get-YardstickIntuneCredential {
    <#
    .SYNOPSIS
    Reads the stored Intune credential from Windows Credential Manager.

    .PARAMETER Target
    Credential Manager target name. Defaults to "Yardstick:IntuneGraph".

    .OUTPUTS
    PSCustomObject with Target, TenantID, ClientID, ClientSecret, SecretExpiresOn
    and LastNotifiedOn, or $null when no credential is stored.
    #>
    [CmdletBinding()]
    param(
        [string]$Target = $Script:DefaultCredentialTarget
    )

    Initialize-YardstickCredentialType

    $userName = $null
    $blob = $null
    if (-not [YardstickCredentialNative]::Read($Target, [ref]$userName, [ref]$blob)) {
        return $null
    }

    $tenantId = $null
    $clientSecret = $null
    $expiresOn = $null
    $lastNotified = $null

    if (-not [string]::IsNullOrWhiteSpace($blob)) {
        $payload = $null
        try {
            $payload = $blob | ConvertFrom-Json -ErrorAction Stop
        } catch {
            $payload = $null
        }

        if ($payload -and $payload.PSObject.Properties['clientSecret']) {
            $tenantId = [string]$payload.tenantId
            $clientSecret = [string]$payload.clientSecret
            $expiresOn = ConvertTo-CredentialDateTime $payload.secretExpiresOn
            $lastNotified = ConvertTo-CredentialDateTime $payload.lastNotifiedOn
        } else {
            # Credential written by hand or by an older layout: the blob is the raw secret.
            $clientSecret = $blob
        }
    }

    return [PSCustomObject]@{
        Target          = $Target
        TenantID        = $tenantId
        ClientID        = if ([string]::IsNullOrWhiteSpace($userName)) { $null } else { $userName.Trim() }
        ClientSecret    = $clientSecret
        SecretExpiresOn = $expiresOn
        LastNotifiedOn  = $lastNotified
    }
}


function Set-YardstickIntuneCredential {
    <#
    .SYNOPSIS
    Writes the Intune credential to Windows Credential Manager.

    .PARAMETER Target
    Credential Manager target name.

    .PARAMETER TenantID
    Entra tenant GUID (or domain name).

    .PARAMETER ClientID
    App registration application (client) ID.

    .PARAMETER ClientSecret
    Client secret value. Accepts a String or SecureString.

    .PARAMETER SecretExpiresOn
    Optional expiration date of the client secret.

    .PARAMETER LastNotifiedOn
    Optional timestamp of the last expiration notification, used to throttle email.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string]$Target = $Script:DefaultCredentialTarget,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantID,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ClientID,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        $ClientSecret,

        [Nullable[datetime]]$SecretExpiresOn,

        [Nullable[datetime]]$LastNotifiedOn
    )

    Initialize-YardstickCredentialType

    $plainSecret = ConvertTo-PlainSecret $ClientSecret
    if ([string]::IsNullOrWhiteSpace($plainSecret)) {
        throw "ClientSecret cannot be empty."
    }

    $payload = [ordered]@{
        tenantId        = $TenantID.Trim()
        clientSecret    = $plainSecret
        secretExpiresOn = if ($SecretExpiresOn) { ([datetime]$SecretExpiresOn).ToUniversalTime().ToString('o') } else { $null }
        lastNotifiedOn  = if ($LastNotifiedOn) { ([datetime]$LastNotifiedOn).ToUniversalTime().ToString('o') } else { $null }
    }

    $blob = $payload | ConvertTo-Json -Compress

    if ($PSCmdlet.ShouldProcess($Target, "Write Yardstick Intune credential to Windows Credential Manager")) {
        [YardstickCredentialNative]::Write($Target, $ClientID.Trim(), $blob, "Yardstick Intune Graph credentials")
        Write-CredentialLog "Stored Intune credentials in Windows Credential Manager under target '$Target'."
    }
}


function Remove-YardstickIntuneCredential {
    <#
    .SYNOPSIS
    Deletes the stored Intune credential from Windows Credential Manager.

    .PARAMETER Target
    Credential Manager target name.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string]$Target = $Script:DefaultCredentialTarget
    )

    Initialize-YardstickCredentialType

    if ($PSCmdlet.ShouldProcess($Target, "Delete Yardstick Intune credential from Windows Credential Manager")) {
        if ([YardstickCredentialNative]::Delete($Target)) {
            Write-CredentialLog "Removed Intune credentials stored under target '$Target'."
            return $true
        }
        Write-CredentialLog "No Intune credentials were stored under target '$Target'."
        return $false
    }
}


function Register-YardstickIntuneCredential {
    <#
    .SYNOPSIS
    Prompts for the Intune app-registration credentials and stores them securely.

    .DESCRIPTION
    Interactive setup used the first time Yardstick runs on a machine, or any
    time the credentials need to be rotated. The client secret is collected as a
    SecureString so it never appears in the console or in PSReadLine history.

    .PARAMETER Target
    Credential Manager target name.

    .PARAMETER TenantID
    Pre-fills the tenant prompt (used when migrating values out of preferences.yaml).

    .PARAMETER ClientID
    Pre-fills the client ID prompt.

    .OUTPUTS
    The stored credential object.
    #>
    [CmdletBinding()]
    param(
        [string]$Target = $Script:DefaultCredentialTarget,
        [string]$TenantID,
        [string]$ClientID
    )

    Write-Host ""
    Write-Host "Yardstick needs Intune (Microsoft Graph) app-registration credentials." -ForegroundColor Cyan
    Write-Host "They will be stored in Windows Credential Manager under target '$Target'," -ForegroundColor Cyan
    Write-Host "encrypted for the current Windows user. Nothing is written to preferences.yaml." -ForegroundColor Cyan
    Write-Host ""

    $tenantPrompt = if ($TenantID) { "Tenant ID [$TenantID]" } else { "Tenant ID" }
    $tenantInput = Read-Host $tenantPrompt
    if ([string]::IsNullOrWhiteSpace($tenantInput)) { $tenantInput = $TenantID }
    if ([string]::IsNullOrWhiteSpace($tenantInput)) { throw "A Tenant ID is required." }

    $clientPrompt = if ($ClientID) { "Client ID [$ClientID]" } else { "Client ID" }
    $clientInput = Read-Host $clientPrompt
    if ([string]::IsNullOrWhiteSpace($clientInput)) { $clientInput = $ClientID }
    if ([string]::IsNullOrWhiteSpace($clientInput)) { throw "A Client ID is required." }

    $secureSecret = Read-Host "Client Secret" -AsSecureString
    if (-not $secureSecret -or $secureSecret.Length -eq 0) { throw "A Client Secret is required." }

    Write-Host "Client secret expiration date (e.g. 2027-01-31). Leave blank to detect it from Graph." -ForegroundColor DarkGray
    $expiryInput = Read-Host "Expires On"
    $expiresOn = $null
    if (-not [string]::IsNullOrWhiteSpace($expiryInput)) {
        $expiresOn = ConvertTo-CredentialDateTime $expiryInput
        if (-not $expiresOn) {
            Write-CredentialLog "Could not parse '$expiryInput' as a date. Expiration will be detected from Graph instead." -Warning
        }
    }

    Set-YardstickIntuneCredential -Target $Target -TenantID $tenantInput -ClientID $clientInput -ClientSecret $secureSecret -SecretExpiresOn $expiresOn
    Write-Host ""

    return Get-YardstickIntuneCredential -Target $Target
}


function Initialize-YardstickIntuneCredential {
    <#
    .SYNOPSIS
    Loads the Intune credentials for a Yardstick run and publishes them as globals.

    .DESCRIPTION
    Resolution order:
      1. Windows Credential Manager (the supported location).
      2. Legacy TenantID / ClientId / ClientSecret keys in preferences.yaml, which
         are migrated into Credential Manager and then flagged for removal.
      3. An interactive prompt, unless -NoPrompt is specified.

    Sets $Global:TenantID, $Global:ClientID and $Global:ClientSecret so
    Connect-YardstickGraph keeps working unchanged.

    .PARAMETER Preferences
    Parsed preferences.yaml hashtable.

    .PARAMETER NoPrompt
    Fail instead of prompting when no credential is stored. Use for scheduled runs.

    .PARAMETER Force
    Always prompt and overwrite the stored credential.

    .OUTPUTS
    The credential object in use.
    #>
    [CmdletBinding()]
    param(
        $Preferences,
        [switch]$NoPrompt,
        [switch]$Force
    )

    $target = Get-YardstickCredentialTarget -Preferences $Preferences

    $legacyTenant = if ($Preferences) { [string]$Preferences.TenantID } else { $null }
    $legacyClient = if ($Preferences) { [string]$Preferences.ClientID } else { $null }
    $legacySecret = if ($Preferences) { [string]$Preferences.ClientSecret } else { $null }
    $hasLegacy = -not ([string]::IsNullOrWhiteSpace($legacyTenant) -or [string]::IsNullOrWhiteSpace($legacyClient) -or [string]::IsNullOrWhiteSpace($legacySecret))

    if ($Force) {
        $credential = Register-YardstickIntuneCredential -Target $target -TenantID $legacyTenant -ClientID $legacyClient
    } else {
        $credential = Get-YardstickIntuneCredential -Target $target
    }

    if ($credential -and $hasLegacy) {
        Write-CredentialLog "preferences.yaml still contains TenantID/ClientId/ClientSecret values. Windows Credential Manager takes precedence - clear those keys from preferences.yaml." -Warning
    }

    if (-not $credential -and $hasLegacy) {
        Write-CredentialLog "Migrating Intune credentials from preferences.yaml into Windows Credential Manager..."
        Set-YardstickIntuneCredential -Target $target -TenantID $legacyTenant -ClientID $legacyClient -ClientSecret $legacySecret
        $credential = Get-YardstickIntuneCredential -Target $target
        Write-CredentialLog "Migration complete. Remove TenantID, ClientId and ClientSecret from preferences.yaml - they are no longer read once credentials are stored." -Warning
    }

    if (-not $credential) {
        if ($NoPrompt) {
            throw "No Intune credentials found in Windows Credential Manager under target '$target'. Run .\Set-YardstickCredential.ps1 to store them before running unattended."
        }
        $credential = Register-YardstickIntuneCredential -Target $target
    }

    # A hand-written credential (raw secret blob, no JSON) has no tenant of its own.
    if ([string]::IsNullOrWhiteSpace($credential.TenantID) -and -not [string]::IsNullOrWhiteSpace($legacyTenant)) {
        $credential.TenantID = $legacyTenant
    }

    $missing = @()
    if ([string]::IsNullOrWhiteSpace($credential.TenantID)) { $missing += 'TenantID' }
    if ([string]::IsNullOrWhiteSpace($credential.ClientID)) { $missing += 'ClientID' }
    if ([string]::IsNullOrWhiteSpace($credential.ClientSecret)) { $missing += 'ClientSecret' }
    if ($missing.Count -gt 0) {
        throw "Stored credential '$target' is incomplete (missing: $($missing -join ', ')). Run .\Set-YardstickCredential.ps1 to re-enter it."
    }

    $Global:TenantID = $credential.TenantID
    $Global:ClientID = $credential.ClientID
    $Global:ClientSecret = $credential.ClientSecret
    $Global:YardstickCredential = $credential

    Write-CredentialLog "Loaded Intune credentials for client $($credential.ClientID) from Windows Credential Manager."
    return $credential
}


function Update-YardstickSecretExpiration {
    <#
    .SYNOPSIS
    Refreshes the stored client secret expiration date from Microsoft Graph.

    .DESCRIPTION
    Reads the app registration's passwordCredentials and matches the one whose
    hint (first three characters) equals the stored secret's. Requires the app
    registration to hold Application.Read.All; if the call is denied the stored
    expiration is left untouched, which is not an error.

    .PARAMETER Credential
    Credential object from Get-YardstickIntuneCredential.

    .OUTPUTS
    The credential object, with SecretExpiresOn refreshed when discovery succeeded.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Credential
    )

    if (-not (Get-Command -Name Invoke-YardstickGraphRequest -ErrorAction SilentlyContinue)) {
        return $Credential
    }

    try {
        $filter = [uri]::EscapeDataString("appId eq '$($Credential.ClientID)'")
        $apps = Invoke-YardstickGraphRequest -ApiVersion 'v1.0' -Resource "applications?`$filter=$filter&`$select=id,displayName,passwordCredentials"
        $app = @($apps) | Select-Object -First 1
        if (-not $app -or -not $app.passwordCredentials) {
            Write-CredentialLog "Graph returned no password credentials for client $($Credential.ClientID); using the stored expiration date."
            return $Credential
        }

        $hint = if ($Credential.ClientSecret.Length -ge 3) { $Credential.ClientSecret.Substring(0, 3) } else { $null }
        $match = $null
        if ($hint) {
            $match = @($app.passwordCredentials | Where-Object { $_.hint -eq $hint }) | Select-Object -First 1
        }
        if (-not $match) {
            # Fall back to the credential that stays valid longest, which is the
            # one an operator most likely just rotated to.
            $match = @($app.passwordCredentials | Sort-Object { [datetime]$_.endDateTime } -Descending) | Select-Object -First 1
        }
        if (-not $match) { return $Credential }

        $expiresOn = ConvertTo-CredentialDateTime $match.endDateTime
        if (-not $expiresOn) { return $Credential }

        if ($Credential.SecretExpiresOn -ne $expiresOn) {
            Set-YardstickIntuneCredential -Target $Credential.Target `
                -TenantID $Credential.TenantID `
                -ClientID $Credential.ClientID `
                -ClientSecret $Credential.ClientSecret `
                -SecretExpiresOn $expiresOn `
                -LastNotifiedOn $Credential.LastNotifiedOn
            $Credential.SecretExpiresOn = $expiresOn
        }
    } catch {
        Write-CredentialLog "Could not read the client secret expiration from Graph ($($_.Exception.Message)). Grant Application.Read.All to the app registration for automatic expiration tracking, or set the date with .\Set-YardstickCredential.ps1."
    }

    return $Credential
}


function Test-YardstickSecretExpiration {
    <#
    .SYNOPSIS
    Warns (and optionally emails) when the Intune client secret expires soon.

    .PARAMETER Credential
    Credential object from Get-YardstickIntuneCredential.

    .PARAMETER Preferences
    Parsed preferences.yaml hashtable. Supplies credentialExpirationWarningDays
    (default 30) and the admin email address.

    .PARAMETER NoEmail
    Log the warning but do not send email.

    .OUTPUTS
    PSCustomObject with ExpiresOn, DaysRemaining, IsExpired, IsExpiring and Notified.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Credential,

        $Preferences,

        [switch]$NoEmail
    )

    $warningDays = $Script:DefaultExpirationWarningDays
    if ($Preferences -and ($Preferences.credentialExpirationWarningDays -as [int]) -gt 0) {
        $warningDays = [int]$Preferences.credentialExpirationWarningDays
    }

    $result = [PSCustomObject]@{
        ExpiresOn     = $Credential.SecretExpiresOn
        DaysRemaining = $null
        IsExpired     = $false
        IsExpiring    = $false
        Notified      = $false
    }

    if (-not $Credential.SecretExpiresOn) {
        Write-CredentialLog "The expiration date of the Intune client secret is unknown. Grant Application.Read.All to the app registration or record the date with .\Set-YardstickCredential.ps1." -Warning
        return $result
    }

    # Whole calendar days, so "expires in 30 days" means the same thing to an
    # operator regardless of the time of day the run happens.
    $daysRemaining = ($Credential.SecretExpiresOn.Date - (Get-Date).Date).Days
    $result.DaysRemaining = $daysRemaining
    $result.IsExpired = $Credential.SecretExpiresOn -lt (Get-Date)
    $result.IsExpiring = $result.IsExpired -or ($daysRemaining -le $warningDays)

    if (-not $result.IsExpiring) {
        Write-CredentialLog "Intune client secret is valid for $daysRemaining more day(s) (expires $($Credential.SecretExpiresOn.ToString('yyyy-MM-dd')))."
        return $result
    }

    $message = if ($result.IsExpired) {
        "The Intune client secret for client $($Credential.ClientID) EXPIRED on $($Credential.SecretExpiresOn.ToString('yyyy-MM-dd')). Yardstick cannot authenticate until it is rotated."
    } else {
        "The Intune client secret for client $($Credential.ClientID) expires in $daysRemaining day(s) on $($Credential.SecretExpiresOn.ToString('yyyy-MM-dd')). Rotate it and run .\Set-YardstickCredential.ps1."
    }
    Write-CredentialLog $message -Warning

    if ($NoEmail) { return $result }

    $intervalHours = $Script:DefaultNotificationIntervalHours
    if ($Preferences -and ($Preferences.credentialExpirationEmailIntervalHours -as [int]) -gt 0) {
        $intervalHours = [int]$Preferences.credentialExpirationEmailIntervalHours
    }
    if ($Credential.LastNotifiedOn -and $Credential.LastNotifiedOn.AddHours($intervalHours) -gt (Get-Date)) {
        Write-CredentialLog "Expiration email already sent within the last $intervalHours hour(s); skipping."
        return $result
    }

    if (Send-YardstickCredentialExpiryEmail -Preferences $Preferences -Credential $Credential -DaysRemaining $daysRemaining -IsExpired:$result.IsExpired) {
        $result.Notified = $true
        $now = Get-Date
        try {
            Set-YardstickIntuneCredential -Target $Credential.Target `
                -TenantID $Credential.TenantID `
                -ClientID $Credential.ClientID `
                -ClientSecret $Credential.ClientSecret `
                -SecretExpiresOn $Credential.SecretExpiresOn `
                -LastNotifiedOn $now
            $Credential.LastNotifiedOn = $now
        } catch {
            Write-CredentialLog "Failed to record the expiration notification timestamp: $($_.Exception.Message)" -Warning
        }
    }

    return $result
}


function Get-YardstickAdminEmailRecipient {
    <#
    .SYNOPSIS
    Resolves the admin notification address, falling back to the report recipient.
    #>
    param($Preferences)

    if (-not $Preferences) { return @() }
    $recipients = if ($Preferences.adminEmailRecipient) { $Preferences.adminEmailRecipient } else { $Preferences.emailRecipient }
    return @($recipients | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
}


function Send-YardstickCredentialExpiryEmail {
    <#
    .SYNOPSIS
    Emails the Yardstick administrator that the Intune client secret is expiring.

    .DESCRIPTION
    Uses the same Outlook COM path as Send-YardstickEmailReport. Sending is gated
    on credentialExpirationEmailEnabled when present, otherwise on
    emailNotificationEnabled.

    .OUTPUTS
    Boolean indicating whether the email was sent.
    #>
    [CmdletBinding()]
    param(
        $Preferences,

        [Parameter(Mandatory = $true)]
        $Credential,

        [int]$DaysRemaining,

        [switch]$IsExpired
    )

    if (-not $Preferences) { return $false }

    $enabled = if ($null -ne $Preferences.credentialExpirationEmailEnabled) {
        [bool]$Preferences.credentialExpirationEmailEnabled
    } else {
        [bool]$Preferences.emailNotificationEnabled
    }
    if (-not $enabled) {
        Write-CredentialLog "Credential expiration email notifications are disabled in preferences."
        return $false
    }

    $recipients = Get-YardstickAdminEmailRecipient -Preferences $Preferences
    if ($recipients.Count -eq 0) {
        Write-CredentialLog "No adminEmailRecipient or emailRecipient configured; cannot send the expiration notice." -Warning
        return $false
    }

    $expiresOn = $Credential.SecretExpiresOn.ToString('yyyy-MM-dd')
    $headline = if ($IsExpired) { "Intune client secret has EXPIRED" } else { "Intune client secret expires in $DaysRemaining day(s)" }
    $accent = if ($IsExpired) { '#dc3545' } else { '#e0a800' }

    $body = @"
<html>
<body style="font-family: Segoe UI, Arial, sans-serif; background-color:#e9ecef; margin:0; padding:20px;">
    <div style="max-width:640px; margin:0 auto; background-color:#f8f9fa; border:1px solid #dee2e6;">
        <div style="background-color:#46C4DD; color:#ffffff; padding:15px;">
            <h2 style="margin:0;">Yardstick Credential Alert</h2>
        </div>
        <div style="padding:20px;">
            <p style="border-left:4px solid $accent; background-color:#ffffff; padding:12px; margin:0 0 16px 0;">
                <strong>$headline</strong>
            </p>
            <table cellpadding="6" cellspacing="0" style="border-collapse:collapse;">
                <tr><td><strong>Tenant ID</strong></td><td>$($Credential.TenantID)</td></tr>
                <tr><td><strong>Client ID</strong></td><td>$($Credential.ClientID)</td></tr>
                <tr><td><strong>Secret expires</strong></td><td>$expiresOn</td></tr>
                <tr><td><strong>Days remaining</strong></td><td>$DaysRemaining</td></tr>
                <tr><td><strong>Machine</strong></td><td>$env:COMPUTERNAME</td></tr>
                <tr><td><strong>Credential target</strong></td><td>$($Credential.Target)</td></tr>
            </table>
            <p>Create a new client secret on the app registration in Entra ID, then store it on the
               Yardstick host by running:</p>
            <pre style="background-color:#ffffff; padding:10px; border:1px solid #dee2e6;">.\Set-YardstickCredential.ps1</pre>
            <p style="color:#6c757d; font-size:0.9em;">Generated automatically by Yardstick.</p>
        </div>
    </div>
</body>
</html>
"@

    try {
        if ((Get-Command -Name Test-OutlookAvailability -ErrorAction SilentlyContinue) -and -not (Test-OutlookAvailability)) {
            Write-CredentialLog "Outlook is not available; cannot send the credential expiration notice." -Warning
            return $false
        }

        $outlook = New-Object -ComObject "Outlook.Application"
        $mail = $outlook.CreateItem(0)
        $mail.To = ($recipients -join "; ")
        $mail.Subject = "Yardstick: $headline"
        if ($Preferences.emailSendFromAddress) {
            $mail.SentOnBehalfOfName = $Preferences.emailSendFromAddress
        }
        $mail.HTMLBody = $body
        $mail.Send()

        Write-CredentialLog "Credential expiration notice sent to $($recipients -join ', ')."
        return $true
    } catch {
        Write-CredentialLog "Failed to send the credential expiration notice: $($_.Exception.Message)" -Warning
        return $false
    } finally {
        $mail = $null
        $outlook = $null
        [System.GC]::Collect()
    }
}


Export-ModuleMember -Function @(
    'Get-YardstickCredentialTarget',
    'Get-YardstickIntuneCredential',
    'Set-YardstickIntuneCredential',
    'Remove-YardstickIntuneCredential',
    'Register-YardstickIntuneCredential',
    'Initialize-YardstickIntuneCredential',
    'Update-YardstickSecretExpiration',
    'Test-YardstickSecretExpiration',
    'Send-YardstickCredentialExpiryEmail',
    'Get-YardstickAdminEmailRecipient'
)
