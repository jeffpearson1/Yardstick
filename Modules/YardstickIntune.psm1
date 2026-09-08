Set-StrictMode -Version 3.0

function Get-YardstickPropertyValue {
    param($InputObject, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [collections.IDictionary]) {
        if ($InputObject.Contains($Name)) { return $InputObject[$Name] }
        return $null
    }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

function ConvertTo-YardstickGraphDate {
    <#
    .SYNOPSIS
    Formats a DateTime for Graph, optionally converting it to UTC first.

    .DESCRIPTION
    Graph carries two different kinds of date in this module and they must not
    be serialized the same way.

    Detection rule dates (createdDate/modifiedDate) name a real instant, so the
    host's offset is meaningful and -AsUtc converts before formatting.

    Install time settings do not. Intune stores the literal wall clock an admin
    typed into the portal and tacks a cosmetic Z onto it; when useLocalTime is
    true the client reads it back in its own timezone, and when useLocalTime is
    false it is already UTC. Either way the number is the answer and converting
    it shifts the deployment by the build host's offset - which is how an
    11:00 PM deployment moved 14 days forward came back as 4:00 AM the next day
    on a UTC-5 host. Those callers leave -AsUtc off.

    .PARAMETER AsUtc
    Convert to UTC before formatting. For values that denote an instant rather
    than a wall clock.
    #>
    param(
        [Parameter(Mandatory)][datetime]$InputObject,
        [switch]$AsUtc
    )
    $value = if ($AsUtc) { $InputObject.ToUniversalTime() } else { $InputObject }
    return $value.ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [cultureinfo]::InvariantCulture)
}

function ConvertTo-YardstickWallClock {
    <#
    .SYNOPSIS
    Reads a Graph date as the wall clock it displays, with no timezone shift.

    .DESCRIPTION
    The same install time setting reaches us as three different types, and left
    alone each reports a different .Hour for one stored value of "23:00:00Z":

      - PowerShell 7 ConvertFrom-Json yields a DateTime of Kind Utc, hour 23.
      - A [datetime] cast of the raw string reinterprets it into the host's
        timezone, hour 18 on a UTC-5 host.
      - A DateTimeOffset carries the offset separately.

    Returning Kind Unspecified in every case keeps the displayed wall clock and
    stops anything downstream from converting it again.
    #>
    param([Parameter(Mandatory)]$InputObject)
    $value = if ($InputObject -is [datetime]) {
        $InputObject
    }
    elseif ($InputObject -is [datetimeoffset]) {
        $InputObject.DateTime
    }
    else {
        # RoundtripKind honours the trailing Z as "this is already the value"
        # instead of treating it as an instant to be moved into local time.
        [datetime]::Parse([string]$InputObject, [cultureinfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::RoundtripKind)
    }
    return [datetime]::SpecifyKind($value, [System.DateTimeKind]::Unspecified)
}

function New-YardstickWin32AppDetectionRuleFile {
    [CmdletBinding(DefaultParameterSetName = 'Existence')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Existence')][switch]$Existence,
        [Parameter(Mandatory, ParameterSetName = 'DateModified')][switch]$DateModified,
        [Parameter(Mandatory, ParameterSetName = 'DateCreated')][switch]$DateCreated,
        [Parameter(Mandatory, ParameterSetName = 'Version')][switch]$Version,
        [Parameter(Mandatory, ParameterSetName = 'Size')][switch]$Size,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$FileOrFolder,
        [bool]$Check32BitOn64System = $false,
        [Parameter(Mandatory, ParameterSetName = 'Existence')][ValidateSet('exists', 'doesNotExist')][string]$DetectionType,
        [Parameter(Mandatory, ParameterSetName = 'DateModified')]
        [Parameter(Mandatory, ParameterSetName = 'DateCreated')]
        [Parameter(Mandatory, ParameterSetName = 'Version')]
        [Parameter(Mandatory, ParameterSetName = 'Size')]
        [ValidateSet('equal', 'notEqual', 'greaterThanOrEqual', 'greaterThan', 'lessThanOrEqual', 'lessThan')]
        [string]$Operator,
        [Parameter(Mandatory, ParameterSetName = 'DateModified')]
        [Parameter(Mandatory, ParameterSetName = 'DateCreated')]
        [datetime]$DateTimeValue,
        [Parameter(Mandatory, ParameterSetName = 'Version')][string]$VersionValue,
        [Parameter(Mandatory, ParameterSetName = 'Size')][string]$SizeInMBValue
    )

    $rule = [ordered]@{
        '@odata.type'          = '#microsoft.graph.win32LobAppFileSystemDetection'
        operator               = if ($PSCmdlet.ParameterSetName -eq 'Existence') { 'notConfigured' } else { $Operator }
        detectionValue         = $null
        path                   = $Path
        fileOrFolderName       = $FileOrFolder
        check32BitOn64System   = $Check32BitOn64System
        detectionType          = $DetectionType
    }
    switch ($PSCmdlet.ParameterSetName) {
        'DateModified' { $rule.detectionType = 'modifiedDate'; $rule.detectionValue = ConvertTo-YardstickGraphDate $DateTimeValue -AsUtc }
        'DateCreated'  { $rule.detectionType = 'createdDate';  $rule.detectionValue = ConvertTo-YardstickGraphDate $DateTimeValue -AsUtc }
        'Version'      { $rule.detectionType = 'version';      $rule.detectionValue = $VersionValue }
        'Size'         { $rule.detectionType = 'sizeInMB';     $rule.detectionValue = $SizeInMBValue }
    }
    return $rule
}

function New-YardstickWin32AppDetectionRuleMsi {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ProductCode,
        [ValidateSet('notConfigured', 'equal', 'notEqual', 'greaterThanOrEqual', 'greaterThan', 'lessThanOrEqual', 'lessThan')]
        [string]$ProductVersionOperator = 'notConfigured',
        [string]$ProductVersion = ''
    )
    return [ordered]@{
        '@odata.type'          = '#microsoft.graph.win32LobAppProductCodeDetection'
        productCode            = $ProductCode
        productVersionOperator = $ProductVersionOperator
        productVersion         = $ProductVersion
    }
}

function New-YardstickWin32AppDetectionRuleRegistry {
    [CmdletBinding(DefaultParameterSetName = 'Existence')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Existence')][switch]$Existence,
        [Parameter(Mandatory, ParameterSetName = 'StringComparison')][switch]$StringComparison,
        [Parameter(Mandatory, ParameterSetName = 'IntegerComparison')][switch]$IntegerComparison,
        [Parameter(Mandatory, ParameterSetName = 'VersionComparison')][switch]$VersionComparison,
        [Parameter(Mandatory)][string]$KeyPath,
        [string]$ValueName = $null,
        [bool]$Check32BitOn64System = $false,
        [Parameter(Mandatory, ParameterSetName = 'Existence')][ValidateSet('exists', 'doesNotExist')][string]$DetectionType,
        [Parameter(Mandatory, ParameterSetName = 'StringComparison')][ValidateSet('equal', 'notEqual')][string]$StringComparisonOperator,
        [Parameter(Mandatory, ParameterSetName = 'IntegerComparison')][ValidateSet('equal', 'notEqual', 'greaterThanOrEqual', 'greaterThan', 'lessThanOrEqual', 'lessThan')][string]$IntegerComparisonOperator,
        [Parameter(Mandatory, ParameterSetName = 'VersionComparison')][ValidateSet('equal', 'notEqual', 'greaterThanOrEqual', 'greaterThan', 'lessThanOrEqual', 'lessThan')][string]$VersionComparisonOperator,
        [Parameter(Mandatory, ParameterSetName = 'StringComparison')][string]$StringComparisonValue,
        [Parameter(Mandatory, ParameterSetName = 'IntegerComparison')][string]$IntegerComparisonValue,
        [Parameter(Mandatory, ParameterSetName = 'VersionComparison')][string]$VersionComparisonValue
    )

    $rule = [ordered]@{
        '@odata.type'        = '#microsoft.graph.win32LobAppRegistryDetection'
        operator             = 'notConfigured'
        detectionValue       = $null
        keyPath              = $KeyPath
        valueName            = $ValueName
        check32BitOn64System = $Check32BitOn64System
        detectionType        = $DetectionType
    }
    switch ($PSCmdlet.ParameterSetName) {
        'StringComparison'  { $rule.operator = $StringComparisonOperator;  $rule.detectionValue = $StringComparisonValue;  $rule.detectionType = 'string' }
        'IntegerComparison' { $rule.operator = $IntegerComparisonOperator; $rule.detectionValue = $IntegerComparisonValue; $rule.detectionType = 'integer' }
        'VersionComparison' { $rule.operator = $VersionComparisonOperator; $rule.detectionValue = $VersionComparisonValue; $rule.detectionType = 'version' }
    }
    return $rule
}

function New-YardstickWin32AppDetectionRuleScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScriptFile,
        [bool]$EnforceSignatureCheck = $false,
        [bool]$RunAs32Bit = $false
    )
    if (-not (Test-Path -LiteralPath $ScriptFile -PathType Leaf)) { throw "Detection script does not exist: $ScriptFile" }
    $content = Get-Content -LiteralPath $ScriptFile -Raw -Encoding UTF8
    return [ordered]@{
        '@odata.type'         = '#microsoft.graph.win32LobAppPowerShellScriptDetection'
        enforceSignatureCheck = $EnforceSignatureCheck
        runAs32Bit             = $RunAs32Bit
        scriptContent          = [convert]::ToBase64String([text.encoding]::UTF8.GetBytes($content))
    }
}

function New-YardstickWin32AppRequirementRule {
    [CmdletBinding()]
    param(
        [ValidateSet('none', 'x64', 'x86', 'arm64', 'x64x86', 'AllWithARM64')]
        [string]$Architecture = 'none',
        [Parameter(Mandatory)]
        [ValidateSet('W10_1607', 'W10_1703', 'W10_1709', 'W10_1803', 'W10_1809', 'W10_1903', 'W10_1909', 'W10_2004', 'W10_20H2', 'W10_21H1', 'W10_21H2', 'W10_22H2', 'W11_21H2', 'W11_22H2', 'W11_23H2', 'W11_24H2')]
        [Alias('MinimumSupportedOperatingSystem')]
        [string]$MinimumSupportedWindowsRelease,
        [int]$MinimumFreeDiskSpaceInMB,
        [int]$MinimumMemoryInMB,
        [int]$MinimumNumberOfProcessors,
        [int]$MinimumCPUSpeedInMHz
    )
    $architectures = @{ none = $null; x64 = 'x64'; x86 = 'x86'; arm64 = 'arm64'; x64x86 = 'x64,x86'; AllWithARM64 = 'x64,x86,arm64' }
    $releases = @{
        W10_1607 = '1607'; W10_1703 = '1703'; W10_1709 = '1709'; W10_1803 = '1803'; W10_1809 = '1809'
        W10_1903 = '1903'; W10_1909 = '1909'; W10_2004 = '2004'; W10_20H2 = '2H20'; W10_21H1 = '21H1'
        W10_21H2 = 'Windows10_21H2'; W10_22H2 = 'Windows10_22H2'; W11_21H2 = 'Windows11_21H2'
        W11_22H2 = 'Windows11_22H2'; W11_23H2 = 'Windows11_23H2'; W11_24H2 = 'Windows11_24H2'
    }
    $rule = [ordered]@{
        allowedArchitectures              = $architectures[$Architecture]
        applicableArchitectures           = 'none'
        minimumSupportedWindowsRelease    = $releases[$MinimumSupportedWindowsRelease]
    }
    foreach ($entry in @(
        @{ Name = 'minimumFreeDiskSpaceInMB'; Value = $MinimumFreeDiskSpaceInMB; Bound = 'MinimumFreeDiskSpaceInMB' }
        @{ Name = 'minimumMemoryInMB'; Value = $MinimumMemoryInMB; Bound = 'MinimumMemoryInMB' }
        @{ Name = 'minimumNumberOfProcessors'; Value = $MinimumNumberOfProcessors; Bound = 'MinimumNumberOfProcessors' }
        @{ Name = 'minimumCpuSpeedInMHz'; Value = $MinimumCPUSpeedInMHz; Bound = 'MinimumCPUSpeedInMHz' }
    )) {
        if ($PSBoundParameters.ContainsKey($entry.Bound)) { $rule[$entry.Name] = $entry.Value }
    }
    return $rule
}

function New-YardstickWin32AppIcon {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$FilePath)
    if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) { throw "Icon does not exist: $FilePath" }
    $extension = [io.path]::GetExtension($FilePath).ToLowerInvariant()
    if ($extension -notin '.png', '.jpg', '.jpeg') { throw "Unsupported icon extension '$extension'." }
    return [ordered]@{
        type  = if ($extension -eq '.png') { 'image/png' } else { 'image/jpeg' }
        value = [convert]::ToBase64String([io.file]::ReadAllBytes((Resolve-Path -LiteralPath $FilePath)))
    }
}

function Get-YardstickWin32App {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'DisplayName')][string]$DisplayName,
        [Parameter(Mandatory, ParameterSetName = 'ID')][string]$ID
    )
    if ($PSCmdlet.ParameterSetName -eq 'ID') {
        return Invoke-YardstickGraphRequest -ApiVersion beta -Resource "deviceAppManagement/mobileApps/$ID"
    }
    $apps = @(Invoke-YardstickGraphRequest -ApiVersion beta -Resource "deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.win32LobApp')")
    if ($PSCmdlet.ParameterSetName -eq 'DisplayName') {
        $apps = @($apps | Where-Object { $_.displayName -like "*$DisplayName*" })
    }
    return $apps
}

function Set-YardstickWin32App {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ID,
        [string]$DisplayName,
        [string]$Description,
        [string]$Publisher,
        [string]$AppVersion,
        [string]$Owner,
        [string]$InstallCommandLine,
        [string]$UninstallCommandLine,
        [bool]$AllowAvailableUninstall
    )
    $body = [ordered]@{ '@odata.type' = '#microsoft.graph.win32LobApp' }
    $map = @{
        DisplayName = 'displayName'; Description = 'description'; Publisher = 'publisher'; AppVersion = 'displayVersion'
        Owner = 'owner'; InstallCommandLine = 'installCommandLine'; UninstallCommandLine = 'uninstallCommandLine'
        AllowAvailableUninstall = 'allowAvailableUninstall'
    }
    foreach ($name in $map.Keys) {
        if ($PSBoundParameters.ContainsKey($name)) { $body[$map[$name]] = $PSBoundParameters[$name] }
    }
    Invoke-YardstickGraphRequest -ApiVersion beta -Method Patch -Resource "deviceAppManagement/mobileApps/$ID" -Body $body | Out-Null
    return Get-YardstickWin32App -ID $ID
}

function Remove-YardstickWin32App {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ID)
    Invoke-YardstickGraphRequest -ApiVersion beta -Method Delete -Resource "deviceAppManagement/mobileApps/$ID" | Out-Null
}

function Get-YardstickWin32AppAssignment {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ID)
    $app = Get-YardstickWin32App -ID $ID
    foreach ($assignment in @(Invoke-YardstickGraphRequest -ApiVersion beta -Resource "deviceAppManagement/mobileApps/$ID/assignments")) {
        $target = Get-YardstickPropertyValue $assignment 'target'
        $settings = Get-YardstickPropertyValue $assignment 'settings'
        $targetType = Get-YardstickPropertyValue $target '@odata.type'
        [pscustomobject]@{
            ID                           = Get-YardstickPropertyValue $assignment 'id'
            Type                         = $targetType
            AppName                      = Get-YardstickPropertyValue $app 'displayName'
            FilterID                     = Get-YardstickPropertyValue $target 'deviceAndAppManagementAssignmentFilterId'
            FilterType                   = Get-YardstickPropertyValue $target 'deviceAndAppManagementAssignmentFilterType'
            GroupID                      = Get-YardstickPropertyValue $target 'groupId'
            GroupName                    = $null
            Intent                       = Get-YardstickPropertyValue $assignment 'intent'
            GroupMode                    = if ($targetType -eq '#microsoft.graph.exclusionGroupAssignmentTarget') { 'Exclude' } elseif ($targetType -eq '#microsoft.graph.groupAssignmentTarget') { 'Include' } else { $null }
            DeliveryOptimizationPriority = Get-YardstickPropertyValue $settings 'deliveryOptimizationPriority'
            Notifications                = Get-YardstickPropertyValue $settings 'notifications'
            RestartSettings              = Get-YardstickPropertyValue $settings 'restartSettings'
            InstallTimeSettings          = Get-YardstickPropertyValue $settings 'installTimeSettings'
        }
    }
}

function New-YardstickAssignmentBody {
    param(
        [Parameter(Mandatory)][string]$TargetType,
        [string]$GroupID,
        [Parameter(Mandatory)][string]$Intent,
        [string]$Notification = 'showAll',
        [datetime]$AvailableTime,
        [datetime]$DeadlineTime,
        [bool]$UseLocalTime = $false,
        [string]$DeliveryOptimizationPriority = 'notConfigured',
        [string]$FilterID,
        [string]$FilterMode = 'Include'
    )
    $target = [ordered]@{
        '@odata.type' = $TargetType
        deviceAndAppManagementAssignmentFilterId = if ($FilterID) { $FilterID } else { $null }
        deviceAndAppManagementAssignmentFilterType = if ($FilterID) { $FilterMode.ToLowerInvariant() } else { 'none' }
    }
    if ($GroupID) { $target.groupId = $GroupID }
    $body = [ordered]@{
        '@odata.type' = '#microsoft.graph.mobileAppAssignment'
        intent        = $Intent
        target        = $target
        settings      = $null
    }
    if ($TargetType -ne '#microsoft.graph.exclusionGroupAssignmentTarget') {
        $body.settings = [ordered]@{
            '@odata.type'                 = '#microsoft.graph.win32LobAppAssignmentSettings'
            notifications                 = $Notification
            restartSettings               = $null
            deliveryOptimizationPriority  = $DeliveryOptimizationPriority
            installTimeSettings           = $null
        }
        if ($PSBoundParameters.ContainsKey('AvailableTime') -or $PSBoundParameters.ContainsKey('DeadlineTime')) {
            # No -AsUtc: Intune stores these as the wall clock the admin sees,
            # not as an instant. See ConvertTo-YardstickGraphDate.
            $body.settings.installTimeSettings = [ordered]@{
                useLocalTime     = $UseLocalTime
                startDateTime    = if ($PSBoundParameters.ContainsKey('AvailableTime')) { ConvertTo-YardstickGraphDate $AvailableTime } else { $null }
                deadlineDateTime = if ($PSBoundParameters.ContainsKey('DeadlineTime')) { ConvertTo-YardstickGraphDate $DeadlineTime } else { $null }
            }
        }
    }
    return $body
}

function Add-YardstickWin32AppAssignment {
    param(
        [Parameter(Mandatory)][string]$ID,
        [Parameter(Mandatory)][string]$TargetType,
        [string]$GroupID,
        [Parameter(Mandatory)][string]$Intent,
        [string]$Notification = 'showAll',
        [datetime]$AvailableTime,
        [datetime]$DeadlineTime,
        [bool]$UseLocalTime = $false,
        [string]$DeliveryOptimizationPriority = 'notConfigured',
        [string]$FilterID,
        [string]$FilterMode = 'Include'
    )
    $bodyParams = @{
        TargetType = $TargetType; GroupID = $GroupID; Intent = $Intent; Notification = $Notification
        UseLocalTime = $UseLocalTime; DeliveryOptimizationPriority = $DeliveryOptimizationPriority
        FilterID = $FilterID; FilterMode = $FilterMode
    }
    if ($PSBoundParameters.ContainsKey('AvailableTime')) { $bodyParams.AvailableTime = $AvailableTime }
    if ($PSBoundParameters.ContainsKey('DeadlineTime')) { $bodyParams.DeadlineTime = $DeadlineTime }
    $body = New-YardstickAssignmentBody @bodyParams
    try {
        return Invoke-YardstickGraphRequest -ApiVersion beta -Method Post -Resource "deviceAppManagement/mobileApps/$ID/assignments" -Body $body
    } catch {
        # Intune answers a duplicate target/intent pair with a BadRequest instead of returning the
        # existing assignment, so resolve it here and keep the add idempotent for callers.
        # The service text lives in ErrorDetails, which is what the error record stringifies to.
        $errorText = @("$_", $_.Exception.Message) -join ' '
        if ($errorText -notmatch 'MobileApp Assignment already exists') { throw }
        $existing = @(Get-YardstickWin32AppAssignmentMatch -ID $ID -TargetType $TargetType -GroupID $GroupID -Intent $Intent)
        if ($existing.Count -eq 0) { throw }
        return $existing[0]
    }
}

function Get-YardstickWin32AppAssignmentMatch {
    param(
        [Parameter(Mandatory)][string]$ID,
        [Parameter(Mandatory)][string]$TargetType,
        [string]$GroupID,
        [Parameter(Mandatory)][string]$Intent
    )
    return @(Invoke-YardstickGraphRequest -ApiVersion beta -Resource "deviceAppManagement/mobileApps/$ID/assignments" | Where-Object {
        $target = Get-YardstickPropertyValue $_ 'target'
        ((Get-YardstickPropertyValue $_ 'intent') -eq $Intent) -and
        ((Get-YardstickPropertyValue $target '@odata.type') -eq $TargetType) -and
        ((-not $GroupID) -or ((Get-YardstickPropertyValue $target 'groupId') -eq $GroupID))
    })
}

function Add-YardstickWin32AppAssignmentGroup {
    [CmdletBinding(DefaultParameterSetName = 'GroupInclude')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'GroupInclude')][switch]$Include,
        [Parameter(Mandatory, ParameterSetName = 'GroupExclude')][switch]$Exclude,
        [Parameter(Mandatory)][string]$ID,
        [Parameter(Mandatory)][string]$GroupID,
        [Parameter(Mandatory)][ValidateSet('required', 'available', 'uninstall')][string]$Intent,
        [string]$Notification = 'showAll',
        [datetime]$AvailableTime,
        [datetime]$DeadlineTime,
        [bool]$UseLocalTime = $false,
        [string]$DeliveryOptimizationPriority = 'notConfigured',
        [string]$FilterName,
        [string]$FilterID,
        [ValidateSet('Include', 'Exclude', 'include', 'exclude')][string]$FilterMode = 'Include'
    )
    if ($FilterName -and -not $FilterID) {
        $filters = @(Invoke-YardstickGraphRequest -ApiVersion beta -Resource 'deviceManagement/assignmentFilters')
        $FilterID = @($filters | Where-Object displayName -eq $FilterName | Select-Object -ExpandProperty id) | Select-Object -First 1
        if (-not $FilterID) { throw "Assignment filter '$FilterName' was not found." }
    }
    $params = @{
        ID = $ID; GroupID = $GroupID; Intent = $Intent; Notification = $Notification; UseLocalTime = $UseLocalTime
        DeliveryOptimizationPriority = $DeliveryOptimizationPriority; FilterID = $FilterID; FilterMode = $FilterMode
        TargetType = if ($Exclude) { '#microsoft.graph.exclusionGroupAssignmentTarget' } else { '#microsoft.graph.groupAssignmentTarget' }
    }
    if ($PSBoundParameters.ContainsKey('AvailableTime')) { $params.AvailableTime = $AvailableTime }
    if ($PSBoundParameters.ContainsKey('DeadlineTime')) { $params.DeadlineTime = $DeadlineTime }
    return Add-YardstickWin32AppAssignment @params
}

function Add-YardstickWin32AppAssignmentAllDevices {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ID,
        [Parameter(Mandatory)][ValidateSet('required', 'available', 'uninstall')][string]$Intent,
        [string]$Notification = 'showAll', [datetime]$AvailableTime, [datetime]$DeadlineTime,
        [bool]$UseLocalTime = $false, [string]$DeliveryOptimizationPriority = 'notConfigured',
        [string]$FilterName, [string]$FilterID, [string]$FilterMode = 'Include'
    )
    if ($FilterName -and -not $FilterID) {
        $FilterID = @(Invoke-YardstickGraphRequest -ApiVersion beta -Resource 'deviceManagement/assignmentFilters' | Where-Object displayName -eq $FilterName | Select-Object -ExpandProperty id) | Select-Object -First 1
        if (-not $FilterID) { throw "Assignment filter '$FilterName' was not found." }
    }
    $params = @{ ID = $ID; TargetType = '#microsoft.graph.allDevicesAssignmentTarget'; Intent = $Intent; Notification = $Notification; UseLocalTime = $UseLocalTime; DeliveryOptimizationPriority = $DeliveryOptimizationPriority; FilterID = $FilterID; FilterMode = $FilterMode }
    if ($PSBoundParameters.ContainsKey('AvailableTime')) { $params.AvailableTime = $AvailableTime }
    if ($PSBoundParameters.ContainsKey('DeadlineTime')) { $params.DeadlineTime = $DeadlineTime }
    return Add-YardstickWin32AppAssignment @params
}

function Add-YardstickWin32AppAssignmentAllUsers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ID,
        [Parameter(Mandatory)][ValidateSet('required', 'available', 'uninstall')][string]$Intent,
        [string]$Notification = 'showAll', [datetime]$AvailableTime, [datetime]$DeadlineTime,
        [bool]$UseLocalTime = $false, [string]$DeliveryOptimizationPriority = 'notConfigured',
        [string]$FilterName, [string]$FilterID, [string]$FilterMode = 'Include'
    )
    if ($FilterName -and -not $FilterID) {
        $FilterID = @(Invoke-YardstickGraphRequest -ApiVersion beta -Resource 'deviceManagement/assignmentFilters' | Where-Object displayName -eq $FilterName | Select-Object -ExpandProperty id) | Select-Object -First 1
        if (-not $FilterID) { throw "Assignment filter '$FilterName' was not found." }
    }
    $params = @{ ID = $ID; TargetType = '#microsoft.graph.allLicensedUsersAssignmentTarget'; Intent = $Intent; Notification = $Notification; UseLocalTime = $UseLocalTime; DeliveryOptimizationPriority = $DeliveryOptimizationPriority; FilterID = $FilterID; FilterMode = $FilterMode }
    if ($PSBoundParameters.ContainsKey('AvailableTime')) { $params.AvailableTime = $AvailableTime }
    if ($PSBoundParameters.ContainsKey('DeadlineTime')) { $params.DeadlineTime = $DeadlineTime }
    return Add-YardstickWin32AppAssignment @params
}

function Remove-YardstickAssignmentsByTarget {
    param([string]$ID, [string]$TargetType, [string]$GroupID)
    foreach ($assignment in @(Invoke-YardstickGraphRequest -ApiVersion beta -Resource "deviceAppManagement/mobileApps/$ID/assignments")) {
        if ($GroupID) {
            if ($assignment.target.groupId -ne $GroupID) { continue }
        } elseif ($assignment.target.'@odata.type' -ne $TargetType) { continue }
        Invoke-YardstickGraphRequest -ApiVersion beta -Method Delete -Resource "deviceAppManagement/mobileApps/$ID/assignments/$($assignment.id)" | Out-Null
    }
}

function Remove-YardstickWin32AppAssignmentGroup { [CmdletBinding()] param([Parameter(Mandatory)][string]$ID, [Parameter(Mandatory)][string]$GroupID) Remove-YardstickAssignmentsByTarget -ID $ID -GroupID $GroupID }
function Remove-YardstickWin32AppAssignmentAllDevices { [CmdletBinding()] param([Parameter(Mandatory)][string]$ID) Remove-YardstickAssignmentsByTarget -ID $ID -TargetType '#microsoft.graph.allDevicesAssignmentTarget' }
function Remove-YardstickWin32AppAssignmentAllUsers { [CmdletBinding()] param([Parameter(Mandatory)][string]$ID) Remove-YardstickAssignmentsByTarget -ID $ID -TargetType '#microsoft.graph.allLicensedUsersAssignmentTarget' }

function Remove-YardstickWin32AppAssignment {
    [CmdletBinding(DefaultParameterSetName = 'ID')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'ID')][string]$ID,
        [Parameter(Mandatory, ParameterSetName = 'DisplayName')][string]$DisplayName
    )
    $apps = if ($PSCmdlet.ParameterSetName -eq 'DisplayName') { @(Get-YardstickWin32App -DisplayName $DisplayName) } else { @(Get-YardstickWin32App -ID $ID) }
    foreach ($app in $apps) {
        foreach ($assignment in @(Invoke-YardstickGraphRequest -ApiVersion beta -Resource "deviceAppManagement/mobileApps/$($app.id)/assignments")) {
            Invoke-YardstickGraphRequest -ApiVersion beta -Method Delete -Resource "deviceAppManagement/mobileApps/$($app.id)/assignments/$($assignment.id)" | Out-Null
        }
    }
}

function Get-YardstickWin32AppRelationship {
    param([Parameter(Mandatory)][string]$ID, [string]$ODataType)
    $relationships = @(Invoke-YardstickGraphRequest -ApiVersion beta -Resource "deviceAppManagement/mobileApps/$ID/relationships")
    if ($ODataType) { return @($relationships | Where-Object { $_.'@odata.type' -eq $ODataType }) }
    return $relationships
}

function Get-YardstickWin32AppDependency { [CmdletBinding()] param([Parameter(Mandatory)][string]$ID) return @(Get-YardstickWin32AppRelationship -ID $ID -ODataType '#microsoft.graph.mobileAppDependency') }
function Get-YardstickWin32AppSupersedence { [CmdletBinding()] param([Parameter(Mandatory)][string]$ID) return @(Get-YardstickWin32AppRelationship -ID $ID -ODataType '#microsoft.graph.mobileAppSupersedence') }

function New-YardstickWin32AppDependency {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ID, [Parameter(Mandatory)][ValidateSet('AutoInstall', 'Detect')][string]$DependencyType)
    if (-not (Get-YardstickWin32App -ID $ID)) { return $null }
    return [ordered]@{ '@odata.type' = '#microsoft.graph.mobileAppDependency'; dependencyType = $DependencyType.ToLowerInvariant(); targetId = $ID }
}

function New-YardstickWin32AppSupersedence {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ID, [Parameter(Mandatory)][ValidateSet('Replace', 'Update')][string]$SupersedenceType)
    if (-not (Get-YardstickWin32App -ID $ID)) { return $null }
    return [ordered]@{ '@odata.type' = '#microsoft.graph.mobileAppSupersedence'; supersedenceType = $SupersedenceType.ToLowerInvariant(); targetId = $ID }
}

function Set-YardstickWin32AppRelationships {
    param([Parameter(Mandatory)][string]$ID, [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Relationships)
    Invoke-YardstickGraphRequest -ApiVersion beta -Method Post -Resource "deviceAppManagement/mobileApps/$ID/updateRelationships" -Body @{ relationships = @($Relationships) } | Out-Null
}

function Add-YardstickWin32AppDependency {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ID, [Parameter(Mandatory)][object[]]$Dependency)
    $supersedence = @(Get-YardstickWin32AppSupersedence -ID $ID | Where-Object { $targetType = Get-YardstickPropertyValue $_ 'targetType'; ($targetType -eq 'child') -or (-not $targetType) })
    Set-YardstickWin32AppRelationships -ID $ID -Relationships (@($Dependency) + $supersedence)
}

function Add-YardstickWin32AppSupersedence {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ID, [Parameter(Mandatory)][object[]]$Supersedence)
    $dependencies = @(Get-YardstickWin32AppDependency -ID $ID | Where-Object { $targetType = Get-YardstickPropertyValue $_ 'targetType'; ($targetType -eq 'child') -or (-not $targetType) })
    Set-YardstickWin32AppRelationships -ID $ID -Relationships (@($Supersedence) + $dependencies)
}

function Remove-YardstickWin32AppDependency {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ID)
    $supersedence = @(Get-YardstickWin32AppSupersedence -ID $ID | Where-Object { $targetType = Get-YardstickPropertyValue $_ 'targetType'; ($targetType -eq 'child') -or (-not $targetType) })
    Set-YardstickWin32AppRelationships -ID $ID -Relationships $supersedence
}

function Remove-YardstickWin32AppSupersedence {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ID)
    $dependencies = @(Get-YardstickWin32AppDependency -ID $ID | Where-Object { $targetType = Get-YardstickPropertyValue $_ 'targetType'; ($targetType -eq 'child') -or (-not $targetType) })
    Set-YardstickWin32AppRelationships -ID $ID -Relationships $dependencies
}

function Resolve-YardstickIntuneWinAppUtil {
    [CmdletBinding()]
    param([string]$ToolsPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Tools'))

    $toolPath = Join-Path $ToolsPath 'IntuneWinAppUtil.exe'
    $expectedHash = 'C1BA45B5CB939E84AF064BB7FF4B38FB3DFE33C8DC1078FD9B157672EAE671F6'
    if (Test-Path -LiteralPath $toolPath -PathType Leaf) {
        $actualHash = (Get-FileHash -LiteralPath $toolPath -Algorithm SHA256).Hash
        if ($actualHash -eq $expectedHash) { return $toolPath }
        throw "IntuneWinAppUtil.exe at '$toolPath' does not match the pinned Microsoft release hash."
    }
    if (-not (Test-Path -LiteralPath $ToolsPath)) { New-Item -ItemType Directory -Path $ToolsPath -Force | Out-Null }
    $downloadPath = "$toolPath.download"
    $downloadUri = 'https://raw.githubusercontent.com/microsoft/Microsoft-Win32-Content-Prep-Tool/1d6cfcbdf8c28edc596337031f74df951f38f718/IntuneWinAppUtil.exe'
    try {
        Invoke-WebRequest -Uri $downloadUri -OutFile $downloadPath -UseBasicParsing -ErrorAction Stop
        $actualHash = (Get-FileHash -LiteralPath $downloadPath -Algorithm SHA256).Hash
        if ($actualHash -ne $expectedHash) { throw "Downloaded IntuneWinAppUtil.exe failed SHA-256 validation." }
        Move-Item -LiteralPath $downloadPath -Destination $toolPath -Force
    } finally {
        if (Test-Path -LiteralPath $downloadPath) { Remove-Item -LiteralPath $downloadPath -Force }
    }
    return $toolPath
}

function New-YardstickWin32AppPackage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [Parameter(Mandatory)][string]$SetupFile,
        [Parameter(Mandatory)][string]$OutputFolder,
        [switch]$Force,
        [string]$IntuneWinAppUtilPath
    )
    if (-not (Test-Path -LiteralPath $SourceFolder -PathType Container)) { throw "Source folder does not exist: $SourceFolder" }
    if (-not (Test-Path -LiteralPath (Join-Path $SourceFolder $SetupFile) -PathType Leaf)) { throw "Setup file '$SetupFile' does not exist in '$SourceFolder'." }
    if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
    if (-not $IntuneWinAppUtilPath) { $IntuneWinAppUtilPath = Resolve-YardstickIntuneWinAppUtil }
    if (-not (Test-Path -LiteralPath $IntuneWinAppUtilPath -PathType Leaf)) { throw "IntuneWinAppUtil.exe does not exist: $IntuneWinAppUtilPath" }

    $packagePath = Join-Path $OutputFolder "$([io.path]::GetFileNameWithoutExtension($SetupFile)).intunewin"
    if ((Test-Path -LiteralPath $packagePath) -and -not $Force) {
        return [pscustomobject]@{ Name = [io.path]::GetFileName($packagePath); Path = $packagePath }
    }
    $arguments = @('-c', ('"{0}"' -f $SourceFolder), '-s', ('"{0}"' -f $SetupFile), '-o', ('"{0}"' -f $OutputFolder), '-q')
    # Give the tool its own console. Shared via -NoNewWindow, its progress-bar cursor
    # writes can leave our console handle unusable and every later Write-Host throws.
    $process = Start-Process -FilePath $IntuneWinAppUtilPath -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
    if ($process.ExitCode -ne 0) { throw "IntuneWinAppUtil.exe exited with code $($process.ExitCode)." }
    if (-not (Test-Path -LiteralPath $packagePath -PathType Leaf)) { throw "IntuneWinAppUtil.exe did not create the expected package: $packagePath" }
    return [pscustomobject]@{ Name = [io.path]::GetFileName($packagePath); Path = $packagePath }
}

function Get-YardstickWin32AppMetadata {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$FilePath)
    if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) { throw "Package does not exist: $FilePath" }
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $archive = [io.compression.zipfile]::OpenRead((Resolve-Path -LiteralPath $FilePath))
    try {
        $entry = @($archive.Entries | Where-Object Name -eq 'detection.xml') | Select-Object -First 1
        if (-not $entry) { throw "Package '$FilePath' does not contain detection.xml." }
        $stream = $entry.Open()
        $reader = [io.streamreader]::new($stream)
        try { return [xml]$reader.ReadToEnd() } finally { $reader.Dispose(); $stream.Dispose() }
    } finally { $archive.Dispose() }
}

function Expand-YardstickWin32AppContent {
    param([Parameter(Mandatory)][string]$FilePath, [Parameter(Mandatory)][string]$FileName)
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $extractRoot = Join-Path ([io.path]::GetTempPath()) ('yardstick-intunewin-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $extractRoot -Force | Out-Null
    $target = Join-Path $extractRoot ([io.path]::GetFileName($FileName))
    $archive = [io.compression.zipfile]::OpenRead((Resolve-Path -LiteralPath $FilePath))
    try {
        $entry = @($archive.Entries | Where-Object Name -eq ([io.path]::GetFileName($FileName))) | Select-Object -First 1
        if (-not $entry) { throw "Package '$FilePath' does not contain encrypted content '$FileName'." }
        [io.compression.zipfileextensions]::ExtractToFile($entry, $target, $true)
    } catch {
        Remove-Item -LiteralPath $extractRoot -Recurse -Force -ErrorAction SilentlyContinue
        throw
    } finally { $archive.Dispose() }
    return $target
}

function Wait-YardstickIntuneFileProcessing {
    param([Parameter(Mandatory)][string]$Resource, [Parameter(Mandatory)][string]$Stage, [int]$TimeoutSeconds = 900, [int]$PollSeconds = 5)
    $deadline = [datetime]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        $file = Invoke-YardstickGraphRequest -ApiVersion beta -Resource $Resource
        switch ($file.uploadState) {
            "$($Stage)Success" { return $file }
            "$($Stage)Failed" { throw "Intune file operation '$Stage' failed." }
            "$($Stage)TimedOut" { throw "Intune file operation '$Stage' timed out in the service." }
        }
        if ([datetime]::UtcNow -ge $deadline) { throw "Timed out waiting for Intune file operation '$Stage'." }
        Start-Sleep -Seconds $PollSeconds
    } while ($true)
}


function Update-YardstickIntuneUploadUri {
    param([Parameter(Mandatory)][string]$FileResource)
    Invoke-YardstickGraphRequest -Method Post -ApiVersion beta -Resource "$FileResource/renewUpload" -Body @{} | Out-Null
    return (Wait-YardstickIntuneFileProcessing -Resource $FileResource -Stage 'azureStorageUriRenewal').azureStorageUri
}

function Invoke-YardstickBlobRequestWithRetry {
    param(
        [Parameter(Mandatory)][ValidateSet('Put')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][byte[]]$Body,
        [hashtable]$Headers,
        [string]$ContentType = 'application/octet-stream',
        [int]$MaximumRetryCount = 5
    )
    for ($attempt = 0; $attempt -le $MaximumRetryCount; $attempt++) {
        try {
            return Invoke-WebRequest -Method $Method -Uri $Uri -Body $Body -Headers $Headers -ContentType $ContentType -UseBasicParsing -ErrorAction Stop
        } catch {
            $status = $null
            if ($_.Exception.Response) {
                try { $status = [int]$_.Exception.Response.StatusCode } catch { $status = $null }
            }
            if ($attempt -ge $MaximumRetryCount -or ($status -and $status -notin 408, 429 -and $status -lt 500)) { throw }
            $delay = [math]::Min(30, [math]::Pow(2, $attempt + 1))
            Start-Sleep -Seconds $delay
        }
    }
}

function Send-YardstickIntuneContentBlob {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$UploadUri,
        [Parameter(Mandatory)][string]$FileResource,
        [int]$ChunkSize = 8MB
    )
    $stream = [io.file]::OpenRead((Resolve-Path -LiteralPath $FilePath))
    $blockIds = [collections.generic.list[string]]::new()
    $uriIssuedAt = [datetime]::UtcNow
    try {
        $index = 0
        $buffer = [byte[]]::new($ChunkSize)
        while (($read = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            if ([datetime]::UtcNow.Subtract($uriIssuedAt).TotalMinutes -ge 7) {
                $UploadUri = Update-YardstickIntuneUploadUri -FileResource $FileResource
                $uriIssuedAt = [datetime]::UtcNow
            }
            $blockId = [convert]::ToBase64String([text.encoding]::ASCII.GetBytes($index.ToString('D6')))
            $blockIds.Add($blockId)
            # Assigning from an if-expression would unroll the array to Object[] and get stringified on the wire.
            [byte[]]$payload = $buffer
            if ($read -ne $buffer.Length) {
                $payload = [byte[]]::new($read)
                [array]::Copy($buffer, $payload, $read)
            }
            $separator = if ($UploadUri.Contains('?')) { '&' } else { '?' }
            $blockUri = "$UploadUri$($separator)comp=block&blockid=$([uri]::EscapeDataString($blockId))"
            Invoke-YardstickBlobRequestWithRetry -Method Put -Uri $blockUri -Body $payload -Headers @{ 'x-ms-blob-type' = 'BlockBlob' } | Out-Null
            $index++
        }
    } finally {
        $stream.Dispose()
    }

    $blockXml = [text.stringbuilder]::new()
    [void]$blockXml.Append('<?xml version="1.0" encoding="utf-8"?><BlockList>')
    foreach ($blockId in $blockIds) { [void]$blockXml.Append("<Latest>$blockId</Latest>") }
    [void]$blockXml.Append('</BlockList>')
    $separator = if ($UploadUri.Contains('?')) { '&' } else { '?' }
    Invoke-YardstickBlobRequestWithRetry -Method Put -Uri "$UploadUri$($separator)comp=blocklist" -Body ([text.encoding]::UTF8.GetBytes($blockXml.ToString())) -Headers @{} -ContentType 'application/xml' | Out-Null
}

function Get-YardstickDefaultReturnCode {
    return @(
        [ordered]@{ returnCode = 0;    type = 'success' }
        [ordered]@{ returnCode = 1707; type = 'success' }
        [ordered]@{ returnCode = 3010; type = 'softReboot' }
        [ordered]@{ returnCode = 1641; type = 'hardReboot' }
        [ordered]@{ returnCode = 1618; type = 'retry' }
    )
}

function Resolve-YardstickScopeTagIds {
    param([string[]]$ScopeTagName)
    if (-not $ScopeTagName) { return @('0') }
    $tags = @(Invoke-YardstickGraphRequest -ApiVersion beta -Resource 'deviceManagement/roleScopeTags')
    $ids = @(
        foreach ($name in $ScopeTagName) {
            $match = @($tags | Where-Object displayName -eq $name)
            if ($match.Count -ne 1) { throw "Could not resolve a unique Intune scope tag named '$name'." }
            $match[0].id
        }
    )
    return $ids
}

function New-YardstickWin32AppBody {
    param(
        [Parameter(Mandatory)][xml]$Metadata,
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][string]$Publisher,
        [Parameter(Mandatory)][string]$InstallCommandLine,
        [Parameter(Mandatory)][string]$UninstallCommandLine,
        [Parameter(Mandatory)]$DetectionRule,
        [Parameter(Mandatory)]$RequirementRule,
        [ValidateSet('system', 'user')][string]$InstallExperience = 'system',
        [ValidateSet('basedOnReturnCode', 'allow', 'suppress', 'force')][string]$RestartBehavior = 'basedOnReturnCode',
        $Icon,
        [string]$AppVersion,
        [string]$Owner,
        [int]$MaximumInstallationTimeInMinutes = 60,
        [switch]$AllowAvailableUninstall,
        [string[]]$ScopeTagName
    )
    $applicationInfo = $Metadata.ApplicationInfo
    # Graph rejects empty strings for these flag enums, so keep unset values as null.
    $allowedArchitectures = Get-YardstickPropertyValue $RequirementRule 'allowedArchitectures'
    if ([string]::IsNullOrWhiteSpace([string]$allowedArchitectures)) { $allowedArchitectures = $null } else { $allowedArchitectures = [string]$allowedArchitectures }
    $applicableArchitectures = Get-YardstickPropertyValue $RequirementRule 'applicableArchitectures'
    if ([string]::IsNullOrWhiteSpace([string]$applicableArchitectures)) { $applicableArchitectures = 'none' } else { $applicableArchitectures = [string]$applicableArchitectures }
    $minimumSupportedWindowsRelease = Get-YardstickPropertyValue $RequirementRule 'minimumSupportedWindowsRelease'
    if ([string]::IsNullOrWhiteSpace([string]$minimumSupportedWindowsRelease)) { $minimumSupportedWindowsRelease = $null } else { $minimumSupportedWindowsRelease = [string]$minimumSupportedWindowsRelease }
    $body = [ordered]@{
        '@odata.type'                    = '#microsoft.graph.win32LobApp'
        displayName                      = $DisplayName
        description                      = $Description
        publisher                        = $Publisher
        displayVersion                   = $AppVersion
        developer                        = ''
        owner                            = $Owner
        notes                            = ''
        informationUrl                   = $null
        privacyInformationUrl            = $null
        isFeatured                       = $false
        fileName                         = [string]$applicationInfo.FileName
        setupFilePath                    = [string]$applicationInfo.SetupFile
        installCommandLine               = $InstallCommandLine
        uninstallCommandLine             = $UninstallCommandLine
        installExperience                = [ordered]@{
            runAsAccount               = $InstallExperience
            deviceRestartBehavior      = $RestartBehavior
            maxRunTimeInMinutes        = $MaximumInstallationTimeInMinutes
        }
        minimumSupportedWindowsRelease  = $minimumSupportedWindowsRelease
        applicableArchitectures         = $applicableArchitectures
        allowedArchitectures            = $allowedArchitectures
        minimumFreeDiskSpaceInMB         = Get-YardstickPropertyValue $RequirementRule 'minimumFreeDiskSpaceInMB'
        minimumMemoryInMB                = Get-YardstickPropertyValue $RequirementRule 'minimumMemoryInMB'
        minimumNumberOfProcessors        = Get-YardstickPropertyValue $RequirementRule 'minimumNumberOfProcessors'
        minimumCpuSpeedInMHz             = Get-YardstickPropertyValue $RequirementRule 'minimumCpuSpeedInMHz'
        msiInformation                   = $null
        runAs32bit                       = $false
        allowAvailableUninstall          = [bool]$AllowAvailableUninstall
        detectionRules                   = @($DetectionRule)
        returnCodes                      = @(Get-YardstickDefaultReturnCode)
        largeIcon                        = if ($Icon) {
            [ordered]@{
                '@odata.type' = '#microsoft.graph.mimeContent'
                type          = if ($Icon -is [collections.IDictionary]) { Get-YardstickPropertyValue $Icon 'type' } else { 'image/png' }
                value         = if ($Icon -is [collections.IDictionary]) { Get-YardstickPropertyValue $Icon 'value' } else { [string]$Icon }
            }
        } else { $null }
        roleScopeTagIds                  = @(Resolve-YardstickScopeTagIds -ScopeTagName $ScopeTagName)
    }
    return $body
}

function Add-YardstickWin32App {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][string]$Publisher,
        [Parameter(Mandatory)][string]$InstallCommandLine,
        [Parameter(Mandatory)][string]$UninstallCommandLine,
        [Parameter(Mandatory)]$DetectionRule,
        [Parameter(Mandatory)]$RequirementRule,
        [ValidateSet('system', 'user')][string]$InstallExperience = 'system',
        [ValidateSet('basedOnReturnCode', 'allow', 'suppress', 'force')][string]$RestartBehavior = 'basedOnReturnCode',
        $Icon,
        [string]$AppVersion,
        [string]$Owner,
        [int]$MaximumInstallationTimeInMinutes = 60,
        [switch]$AllowAvailableUninstall,
        [string[]]$ScopeTagName
    )
    $metadata = Get-YardstickWin32AppMetadata -FilePath $FilePath
    $applicationInfo = $metadata.ApplicationInfo
    $app = $null
    $expandedContent = $null
    try {
        $bodyParameters = @{}
        foreach ($key in $PSBoundParameters.Keys) {
            if ($key -ne 'FilePath') { $bodyParameters[$key] = $PSBoundParameters[$key] }
        }
        $body = New-YardstickWin32AppBody @bodyParameters -Metadata $metadata
        $app = Invoke-YardstickGraphRequest -Method Post -ApiVersion beta -Resource 'deviceAppManagement/mobileApps' -Body $body
        $contentVersion = Invoke-YardstickGraphRequest -Method Post -ApiVersion beta -Resource "deviceAppManagement/mobileApps/$($app.id)/microsoft.graph.win32LobApp/contentVersions" -Body @{}
        $expandedContent = Expand-YardstickWin32AppContent -FilePath $FilePath -FileName ([string]$applicationInfo.FileName)
        $encryptedSize = (Get-Item -LiteralPath $expandedContent).Length
        $fileBody = [ordered]@{
            '@odata.type' = '#microsoft.graph.mobileAppContentFile'
            name          = [io.path]::GetFileName($FilePath)
            size          = [int64]$applicationInfo.UnencryptedContentSize
            sizeEncrypted = [int64]$encryptedSize
            manifest      = $null
            isDependency  = $false
        }
        $fileResource = "deviceAppManagement/mobileApps/$($app.id)/microsoft.graph.win32LobApp/contentVersions/$($contentVersion.id)/files"
        $contentFile = Invoke-YardstickGraphRequest -Method Post -ApiVersion beta -Resource $fileResource -Body $fileBody
        $contentFileResource = "$fileResource/$($contentFile.id)"
        $contentFile = Wait-YardstickIntuneFileProcessing -Resource $contentFileResource -Stage 'azureStorageUriRequest'
        Send-YardstickIntuneContentBlob -FilePath $expandedContent -UploadUri $contentFile.azureStorageUri -FileResource $contentFileResource
        $encryptionInfo = $applicationInfo.EncryptionInfo
        $commitBody = [ordered]@{
            fileEncryptionInfo = [ordered]@{
                encryptionKey        = [string]$encryptionInfo.EncryptionKey
                macKey               = [string]$encryptionInfo.MacKey
                initializationVector = [string]$encryptionInfo.InitializationVector
                mac                  = [string]$encryptionInfo.Mac
                profileIdentifier    = [string]$encryptionInfo.ProfileIdentifier
                fileDigest           = [string]$encryptionInfo.FileDigest
                fileDigestAlgorithm  = [string]$encryptionInfo.FileDigestAlgorithm
            }
        }
        Invoke-YardstickGraphRequest -Method Post -ApiVersion beta -Resource "$contentFileResource/commit" -Body $commitBody | Out-Null
        Wait-YardstickIntuneFileProcessing -Resource $contentFileResource -Stage 'commitFile' | Out-Null
        Invoke-YardstickGraphRequest -Method Patch -ApiVersion beta -Resource "deviceAppManagement/mobileApps/$($app.id)" -Body ([ordered]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; committedContentVersion = [string]$contentVersion.id }) | Out-Null
        return Get-YardstickWin32App -ID $app.id
    } catch {
        if ($app -and $app.id) {
            try { Remove-YardstickWin32App -ID $app.id -ErrorAction Stop } catch { Write-Warning "Failed to remove incomplete Intune app $($app.id): $($_.Exception.Message)" }
        }
        throw
    } finally {
        if ($expandedContent) {
            $extractRoot = Split-Path -Parent $expandedContent
            if (Test-Path -LiteralPath $extractRoot) { Remove-Item -LiteralPath $extractRoot -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
}

Export-ModuleMember -Function @(
    # Exported for YardstickSupport, which rebases install times across apps.
    'ConvertTo-YardstickWallClock',
    'New-YardstickWin32AppDetectionRuleFile',
    'New-YardstickWin32AppDetectionRuleMsi',
    'New-YardstickWin32AppDetectionRuleRegistry',
    'New-YardstickWin32AppDetectionRuleScript',
    'New-YardstickWin32AppRequirementRule',
    'New-YardstickWin32AppIcon',
    'New-YardstickWin32AppPackage',
    'Get-YardstickWin32App',
    'Add-YardstickWin32App',
    'Set-YardstickWin32App',
    'Remove-YardstickWin32App',
    'Get-YardstickWin32AppAssignment',
    'Add-YardstickWin32AppAssignment',
    'Add-YardstickWin32AppAssignmentGroup',
    'Add-YardstickWin32AppAssignmentAllDevices',
    'Add-YardstickWin32AppAssignmentAllUsers',
    'Remove-YardstickWin32AppAssignment',
    'Remove-YardstickWin32AppAssignmentGroup',
    'Remove-YardstickWin32AppAssignmentAllDevices',
    'Remove-YardstickWin32AppAssignmentAllUsers',
    'Get-YardstickWin32AppRelationship',
    'Get-YardstickWin32AppDependency',
    'Get-YardstickWin32AppSupersedence',
    'New-YardstickWin32AppDependency',
    'New-YardstickWin32AppSupersedence',
    'Add-YardstickWin32AppDependency',
    'Add-YardstickWin32AppSupersedence',
    'Remove-YardstickWin32AppDependency',
    'Remove-YardstickWin32AppSupersedence'
)
