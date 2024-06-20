# Useful Functions

# Write-Log
# Prints both to the console and also to the log
# Param: [String] $ContentToWriteToLog, [Switch] $Init (Erase log and print big timestamp for new script run)
function Write-Log {
    param(
        [String]$Content,
        [Switch]$Init
    )
    Push-Location $LOG_LOCATION
    if($Init) {
        # if (Test-Path $LOG_LOCATION\$LOG_FILE) {
        #     Remove-Item $LOG_LOCATION\$LOG_FILE -Force
        # }
        Write-Output "#######################################################" | Out-File $LOG_FILE -Append
        Write-Output "LOGGING STARTED AT $(Get-Date -Format "MM/dd/yyyy HH:mm:ss")" | Out-File $LOG_FILE -Append
        Write-Output "#######################################################" | Out-File $LOG_FILE -Append
    }
    if ($Content) {
        $Content = "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss") - $Content"
        Write-Output $Content | Out-File $LOG_FILE -Append
        Write-Output $Content
    }
    Pop-Location
}


# ArrayToString
# Converts an array to a string
function ArrayToString() {
    param (
        [Array] $array
    )
    $arrayString = "@("
    foreach ($value in $array) {
        $arrayString = "$($arrayString)$([String]$value),"
    }
    $arrayString = $arrayString.TrimEnd(",")
    $arrayString = "$arrayString)"
    return [String]$arrayString
}

# Get-RedirectedUrl
# Follows redirects and returns the actual URL of a resource
# Param: [String] URL
# Return: [String] The last URL that is linked to that does not redirect
function Get-RedirectedUrl() {
    param (
        [Parameter(Mandatory=$true)]
        [String]$URL
    )

    $request = [System.Net.WebRequest]::Create($url)
    $request.AllowAutoRedirect=$false
    $response=$request.GetResponse()

    If ($response.StatusCode -eq "Found")
    {
        return $response.GetResponseHeader("Location")
    }
}


# Get-MsiProductCode
# Returns the product code of an MSI file in the current directory only
# Param: [String] $filePath
# Return: [String] The MSI Product Code
function Get-MsiProductCode() {
    param (
        [Parameter(Mandatory=$true)]
        [String]$filePath
    )
    # Read property from MSI database
    $windowsInstallerObject = New-Object -ComObject WindowsInstaller.Installer
    $msiDatabase = $windowsInstallerObject.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $windowsInstallerObject, @($filePath, 0))
    $query = "SELECT Value FROM Property WHERE Property = 'ProductCode'"
    $view = $msiDatabase.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $msiDatabase, ($query))
    $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)
    $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
    $value = $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1)
    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($windowsInstallerObject) 
    return [String]$value
}

# Connect-AutoMSIntuneGraph
# Automatically connects to the Microsoft Graph API using the Intune module
# Param: None
# Returns: None
function Connect-AutoMSIntuneGraph() {
    # Check if the current token is invalid
    if (-not $Global:Token) {
        Write-Log "Getting an Intune Graph Client APItoken..."
        $Global:Token = Connect-MSIntuneGraph -TenantID $TENANT_ID -ClientID $CLIENT_ID -ClientSecret $CLIENT_SECRET
    }
    elseif ($Global:Token.ExpiresOn.ToLocalTime() -lt (Get-Date)) {
        # If not, get a new token
        Write-Log "Token is expired. Refreshing token..."
        $Global:Token = Connect-MSIntuneGraph -TenantID $TENANT_ID -ClientID $CLIENT_ID -ClientSecret $CLIENT_SECRET
        Write-Log "Token refreshed. New Token Expires at: $($Global:Token.ExpiresOn.ToLocalTime())"
    }
    elseif ($Global:Token.ExpiresOn.AddMinutes(-30).ToLocalTime() -lt (Get-Date)) {
        # For whatever reason, this API stops working 10 minutes before a token refresh
        # Set at 30 minutes in case we are uploading large files. 
        Write-Log "Token expires soon - Refreshing token..."
        # Required to force a refresh
        Clear-MsalTokenCache
        $Global:Token = Connect-MSIntuneGraph -TenantID $TENANT_ID -ClientID $CLIENT_ID -ClientSecret $CLIENT_SECRET
        Write-Log "Token refreshed. New Token Expires at: $($Global:Token.ExpiresOn.ToLocalTime())"
    }
    else {
        Write-Log "Token is still valid. Skipping token refresh."
    } 
}

# MOVE-ASSIGNMENTS
# Gets all assignments for an application and adds them to a target (To)
# Removes successfully moved assignments from the (From) application
# Param: [Application]$From, [Application]$To
# Return: None
function Move-Assignments {
    param(
        [Parameter(Mandatory, Position=0)]
        [System.Object] $From,
        [Parameter(Mandatory, Position=1)]
        [System.Object] $To
    )
    Write-Log "DEBUG: FROM.ID - $($From.id)"
    Write-Log "DEBUG: TO.ID - $($To.id)"
    $FromAssignments = Get-IntuneWin32AppAssignment -Id $From.id
    $FromAvailable = ($FromAssignments | Where-Object intent -eq "available").groupId
    $FromRequired = ($FromAssignments | Where-Object intent -eq "required").groupId
    if ($FromAvailable) {
        foreach ($groupId in $FromAvailable) {
            $maxRetries = 3
            $try = 0
            $successfullyAdded = $false
            while (!$successfullyAdded -and ($try -lt $maxRetries)) {
                Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $groupId -Intent "available" -Notification "hideAll" | Out-Null
                # Check that it worked
                $AssignedGroups = Get-IntuneWin32AppAssignment -ID $To.id
                if($AssignedGroups | Where-Object GroupID -eq $groupID) {
                    $successfullyAdded = $true
                    Write-Log "$GroupID (Available) added successfully to $($To.id)"
                    Write-Log "Removing $groupID available assignment from $($From.id)"
                    Remove-IntuneWin32AppAssignmentGroup -ID $From.id -GroupID $groupId | Out-Null
                }
                else {
                    # Wait 1 second and retry
                    Start-Sleep -Seconds 1
                }
            }
            
            if (!$successfullyAdded) {
                Write-Log "ERROR: Cannot add available app assignment to $($To.id) for group $groupID. Will not remove original assignment."
            }
        }
        
    }
    if ($FromRequired) {
        foreach ($groupId in $FromRequired) {
            $maxRetries = 3
            $try = 0
            $successfullyAdded = $false
            while (!$successfullyAdded -and ($try -lt $maxRetries)) {
                Add-IntuneWin32AppAssignmentGroup -Include -ID $To.id -GroupID $groupId -Intent "required" -Notification "hideAll" | Out-Null
                # Check that it worked
                $AssignedGroups = Get-IntuneWin32AppAssignment -ID $To.id
                if($AssignedGroups | Where-Object GroupID -eq $groupID) {
                    $successfullyAdded = $true
                    Write-Log "$GroupID (Required) added successfully to $($To.id)"
                    Write-Log "Removing $groupID required assignment from $($From.id)"
                    Remove-IntuneWin32AppAssignmentGroup -ID $From.id -GroupID $groupId | Out-Null
                }
                else {
                    # Wait 1 second and retry
                    Start-Sleep -Seconds 1
                }
            }
            
            if (!$successfullyAdded) {
                Write-Log "ERROR: Cannot add required app assignment to $($To.id) for group $groupID. Will not remove original assignment."
            }
        }
    }
}


# Get-SameAppAllVersions
# Returns all versions of an app sorted from newest to oldest
# Accounts for edge cases where an application name might be similar to others, i.e. Mozilla Firefox vs. Mozilla Firefox ESR
# Param: [String] DisplayName
# Return: @(PSCustomObject) 
function Get-SameAppAllVersions($DisplayName) {
    $AllSimilarApps = Get-IntuneWin32App -DisplayName "$DisplayName*"
    $sortable = ($AllSimilarApps | Where-Object {($_.DisplayName -eq $DisplayName) -or ($_.DisplayName -like "$DisplayName (*")})
    return $sortable | Sort-Object {$($($_.displayVersion -replace "[A-Za-z]", "0").split(".") | ForEach-Object {'{0:d8}' -f [int]$_}) -join ''} -Descending
}
