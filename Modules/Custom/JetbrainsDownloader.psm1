$JetbrainsXMLURL = "https://www.jetbrains.com/updates/updates.xml"
$JetbrainsXML = [xml](Invoke-WebRequest -Uri $JetbrainsXMLURL).Content


function Get-JetbrainsAppLatestVersion {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("AppCode", "CLion", "DataGrip", "GoLand", "IntelliJ IDEA", "PhpStorm", "PyCharm", "Rider", "RubyMine", "RustRover", "WebStorm")]
        [string]$ProductName,
        [Parameter(Mandatory=$true)]
        [ValidateSet("release", "eap")]
        [string]$Channel
    )

    $ProductId = Get-JetbrainsAppProductId -ProductName $ProductName -Channel $Channel
    $App = $JetbrainsXML.products.product | Where-Object { $_.name -eq $ProductName } | Select-Object -First 1
    $App = $App.channel | Where-Object { $_.id -eq $ProductId } | Select-Object -First 1
    $Version = ($App.build | Select-Object -First 1).version
    return $Version

}

function Get-JetbrainsAppProductId {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("AppCode", "CLion", "DataGrip", "GoLand", "IntelliJ IDEA", "PhpStorm", "PyCharm", "Rider", "RubyMine", "RustRover", "WebStorm")]
        [string]$ProductName,
        [Parameter(Mandatory=$true)]
        [ValidateSet("release", "eap")]
        [string]$Channel
    )
    if ($Channel -eq "release") {
        $ProductId = switch ($ProductName) {
            "AppCode" { "OC-RELEASE-licensing-RELEASE" }
            "CLion" { "CL-RELEASE-licensing-RELEASE" }
            "DataGrip" { "DB-RELEASE-licensing-RELEASE" }
            "GoLand" { "GO-RELEASE-licensing-RELEASE" }
            "IntelliJ IDEA" { "IC-IU-RELEASE-licensing-RELEASE" }
            "PhpStorm" { "PS-RELEASE-licensing-RELEASE" }
            "PyCharm" { "PC-PY-RELEASE-licensing-RELEASE" }
            "Rider" { "RD-RELEASE-licensing-RELEASE" }
            "RubyMine" { "RM-RELEASE-licensing-RELEASE" }
            "RustRover" { "RR-RELEASE-licensing-RELEASE" }
            "WebStorm" { "WS-RELEASE-licensing-RELEASE" }
            Default { "None" }
        }
    }
    else {
        $ProductId = switch ($ProductName) {
            "CLion" { "CL-EAP-licensing-EAP" }
            "DataGrip" { "DB-EAP-licensing-EAP" }
            "GoLand" { "GO-EAP-licensing-EAP" }
            "IntelliJ IDEA" { "IC-IU-EAP-licensing-EAP" }
            "PhpStorm" { "PS-EAP-licensing-EAP" }
            "PyCharm" { "PC-EAP-licensing-EAP" }
            "Rider" { "RD-EAP-licensing-EAP" }
            "RubyMine" { "RM-EAP-licensing-EAP" }
            "RustRover" { "RR-EAP-licensing-EAP" }
            "WebStorm" { "WS-EAP-licensing-EAP" }
            Default { "None" }
        }
    }
    if ($ProductId -eq "None") {
        Write-Error "Product ID not found for $ProductName in $Channel channel."
        return $null
    }
    return $ProductId
}

function Get-JetbrainsAppDownloadLink {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [Parameter(ParameterSetName="Latest")]
        [Parameter(ParameterSetName="SpecificVersion")]
        [ValidateSet("AppCode", "CLion", "DataGrip", "GoLand", "IntelliJ IDEA", "PhpStorm", "PyCharm", "Rider", "RubyMine", "RustRover", "WebStorm")]
        [string]$ProductName,
        [Parameter(Mandatory=$true)]
        [Parameter(ParameterSetName="Latest")]
        [Parameter(ParameterSetName="SpecificVersion")]
        [ValidateSet("release", "eap")]
        [string]$Channel,
        [Parameter(ParameterSetName="Latest")]
        [Switch]$Latest,
        [Parameter(ParameterSetName="SpecificVersion")]
        [string]$Version,
        [Parameter()]
        [ValidateSet("x64", "arm64")]
        [string]$Architecture = "x64",
        [Parameter()]
        [switch]$ValidateUrl
    )
    
    if ($Latest) {
        $Version = Get-JetbrainsAppLatestVersion -ProductName $ProductName -Channel $Channel
    }
    
    if (-not $Version) {
        Write-Error "Version is required. Use -Latest or provide -Version parameter."
        return $null
    }
    
    # Map product names to download URL components
    # Path = folder on download.jetbrains.com, Prefix = filename prefix
    $DownloadMap = @{
        "AppCode"       = @{ Path = "objc"; Prefix = "AppCode" }
        "CLion"         = @{ Path = "cpp"; Prefix = "CLion" }
        "DataGrip"      = @{ Path = "datagrip"; Prefix = "datagrip" }
        "GoLand"        = @{ Path = "go"; Prefix = "goland" }
        "IntelliJ IDEA" = @{ Path = "idea"; Prefix = "ideaIU" }
        "PhpStorm"      = @{ Path = "webide"; Prefix = "PhpStorm" }
        "PyCharm"       = @{ Path = "python"; Prefix = "pycharm-professional" }
        "Rider"         = @{ Path = "rider"; Prefix = "JetBrains.Rider" }
        "RubyMine"      = @{ Path = "ruby"; Prefix = "RubyMine" }
        "RustRover"     = @{ Path = "rustrover"; Prefix = "RustRover" }
        "WebStorm"      = @{ Path = "webstorm"; Prefix = "WebStorm" }
    }
    
    $ProductInfo = $DownloadMap[$ProductName]
    if (-not $ProductInfo) {
        Write-Error "Download mapping not found for $ProductName"
        return $null
    }
    
    # Construct download URL based on architecture
    $BaseUrl = "https://download.jetbrains.com/$($ProductInfo.Path)"
    
    if ($Architecture -eq "arm64") {
        $FileName = "$($ProductInfo.Prefix)-$Version-aarch64.exe"
    }
    else {
        $FileName = "$($ProductInfo.Prefix)-$Version.exe"
    }
    
    $DownloadUrl = "$BaseUrl/$FileName"
    
    # Optionally validate the URL exists
    if ($ValidateUrl) {
        $isValid = Test-JetbrainsDownloadUrl -Url $DownloadUrl
        if (-not $isValid) {
            Write-Error "Download URL validation failed for: $DownloadUrl"
            return $null
        }
    }
    
    return [PSCustomObject]@{
        ProductName  = $ProductName
        Version      = $Version
        FileName     = $FileName
        DownloadUrl  = $DownloadUrl
        Channel      = $Channel
        Architecture = $Architecture
    }
}

function Test-JetbrainsDownloadUrl {
    <#
    .SYNOPSIS
        Validates that a JetBrains download URL is accessible.
    .DESCRIPTION
        Performs a HEAD request to verify the download URL returns a successful status code.
    .PARAMETER Url
        The download URL to validate.
    .EXAMPLE
        Test-JetbrainsDownloadUrl -Url "https://download.jetbrains.com/python/pycharm-professional-2025.3.2.exe"
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Url
    )
    
    try {
        $response = Invoke-WebRequest -Uri $Url -Method Head -UseBasicParsing -ErrorAction Stop
        return $response.StatusCode -eq 200
    }
    catch {
        Write-Verbose "URL validation failed for $Url : $($_.Exception.Message)"
        return $false
    }
}