# Uses Selenium in a similar manner to AdobeDownloader.psm1 to retrieve the latest versions of Adobe products and saves them to a JSON file.

<#
.SYNOPSIS
Initializes the Selenium WebDriver for browser automation.
#>
function InitializeBrowser() {
    $consoleURL = "https://adminconsole.adobe.com"
    
    try {
        Write-Log "Starting Firefox WebDriver..."
        $driver = Start-SeDriver -Browser "Firefox" -StartURL $consoleURL
        Start-Sleep -Seconds 15
        return $driver
    }
    catch {
        throw "Cannot start Firefox with Selenium: $_"
    }
}

<#
.SYNOPSIS
Handles authentication to Adobe Admin Console.
#>
function AuthenticateToAdobe([object]$Driver, [switch]$SSO) {
    # Load credentials
    try {
        $credentials = Get-StoredCredential -Target "adobe"
        if (-not $credentials) {
            throw "No credentials found for target 'adobe'"
        }
    }
    catch {
        throw "Cannot retrieve stored credentials with target 'adobe': $_"
    }

    # Enter username
    try {
        Write-Log "Entering username..."
        $element = Get-SeElement -By ID "EmailPage-EmailField"
        Invoke-SeKeys -Element $element -Keys $credentials.Username
        $button = Get-SeElement -By XPath '//*[@class="EmailPage__buttons"]/button'
        $button.click()
        Start-Sleep -Seconds 15
    }
    catch {
        throw "Cannot complete username entry: $_"
    }

    # Handle authentication (SSO or password)
    if ($SSO) {
        try {
            Write-Log "Authenticating via SSO..."
            Invoke-SSOAutoSignIn -Target "adminconsole.adobe.com" -Driver $Driver
        }
        catch {
            throw "SSO authentication failed: $_"
        }
    }
    else {
        try {
            Write-Log "Entering password..."
            $element = Get-SeElement -By ID "PasswordPage-PasswordField"
            $password = (New-Object PSCredential 0, $credentials.Password).GetNetworkCredential().Password
            Invoke-SeKeys -Element $element -Keys $password
            $button = Get-SeElement -By CSSSelector ".spectrum-Button"
            $button.click()
            Start-Sleep -Seconds 10
        }
        catch {
            throw "Password authentication failed: $_"
        }
    }
}

<#
.SYNOPSIS
Navigates to the Packages section in Adobe Admin Console.
#>
function NavigateToPackages([object]$Driver) {
    try {
        # Wait for page to load completely
        Write-Log "Waiting for Adobe Admin Console to load..."
        $timeout = 60 # seconds
        $elapsed = 0
        while (-not (Get-SeElement -By LinkText "Overview" -ErrorAction SilentlyContinue) -and $elapsed -lt $timeout) {
            Start-Sleep -Seconds 3
            $elapsed += 3
        }
        
        if ($elapsed -ge $timeout) {
            throw "Timeout waiting for Adobe Admin Console to load"
        }
        
        Write-Log "Navigating to Packages tab..."
        $packagesLink = Get-SeElement -By LinkText "Packages"
        $packagesLink.click()
        Start-Sleep -Seconds 5
    }
    catch {
        throw "Cannot navigate to Packages tab: $_"
    }
}

<#
.SYNOPSIS
Helper method to click the Next button with logging.
#>
function ClickNextButton([string]$Context) {
    Write-Log "Clicking Next button - $Context"
    $nextButton = Get-SeElement -By Xpath '//*[@id="react-aria2916278133-:r6v:"]' -ErrorAction SilentlyContinue
    if ($nextButton) {
        $nextButton.click()
        return
    }
    $nextButton = Get-SeElement -By Xpath '//*[@id="react-aria2916278133-:r8d:"]' -ErrorAction SilentlyContinue
    if ($nextButton) {
        $nextButton.click()
        Start-Sleep -Seconds 2
        return
    }
}

<#
.SYNOPSIS
Helper method to select platform (Windows 64-bit).
#>
function SelectPlatform([string]$Platform) {
    Write-Log "Selecting platform: $Platform"
    $platformDropdown = Get-SeElement -By XPath "//*[text()='Select platform']"
    $platformDropdown.click()
    Start-Sleep -Seconds 2
    
    $platformOption = Get-SeElement -By XPath "//*[text()='$Platform']"
    $platformOption.click()
    Start-Sleep -Seconds 2
}

<#
.SYNOPSIS
Navigates to the package list under 'Create a Package'
#>
function NavigateToPackageList() {
    try {
            # Click create package button
            Write-Log "Waiting for 'Create Package' button..."
            $timeout = 30
            $elapsed = 0
            while (-not (Get-SeElement -By CSSSelector "button.Dniwja_spectrum-Button:nth-child(2)" -ErrorAction SilentlyContinue) -and $elapsed -lt $timeout) {
                Start-Sleep -Seconds 3
                $elapsed += 3
            }
            
            if ($elapsed -ge $timeout) {
                throw "Timeout waiting for Create Package button"
            }
            
            Write-Log "Clicking 'Create Package'..."
            $createButton = Get-SeElement -By CSSSelector "button.Dniwja_spectrum-Button:nth-child(2)"
            $createButton.click()
            Start-Sleep -Seconds 5

            # Named user is fine for this, so just click next
            ClickNextButton("Named user licensing")

            # Select managed package
            Write-Log "Selecting managed package option..."
            $managedInput = Get-SeElement -By XPath "//*[text()='Managed package']/../following-sibling::*/input"
            $managedInput.click()
            Start-Sleep -Seconds 2
            ClickNextButton("Package type selection")
            SelectPlatform("Windows (64-bit)")
            
    }
    catch {
        throw "Cannot navigate to package list: $_"
    }
}

<#
.SYNOPSIS
Saves all versions of available Adobe products
#>
function SaveProductVersions() {
    $list = (Get-SeElement -By Xpath '//*[@class="spectrum-Accordion"]/div/div').Text
    # Split $list using parentheses as the delimiter, and save the results as a JSON file
    $products = @{}
    foreach ($item in $list) {
        $components = $item -split '\s*\(\s*'
        if ($components.Count -eq 2) {
            $products.Add($components[0], $components[1].trim(')'))
        }
    }
    $products | ConvertTo-Json | Set-Content -Path "$Global:SCRIPT_ROOT\AdobeProductVersions.json"
}



# Run the script
Import-Module "$Global:SCRIPT_ROOT\Modules\Custom\AdobeDownloader.psm1"
$Driver = InitializeBrowser
AuthenticateToAdobe -Driver $Driver -SSO
NavigateToPackages -Driver $Driver
NavigateToPackageList
SaveProductVersions
$Driver.Quit()

