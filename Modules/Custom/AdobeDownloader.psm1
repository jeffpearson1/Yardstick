<#
.SYNOPSIS
AdobeApplication class for automating Adobe Creative Cloud package creation and download.

.DESCRIPTION
This class provides functionality to create and download Adobe Creative Cloud packages
through the Adobe Admin Console using Selenium WebDriver automation.
#>

class AdobeApplication {
    # Basic application properties
    [string]$Name
    [string]$ID
    [string]$URL
    [string]$VersionMatchStringRegex
    [string]$Version
    [string]$VersionLocation
    [string]$FileDetectionPath
    [string]$PackageName
    
    # Installation configuration
    [string]$InstallScript = 'setup.exe --silent --ADOBEINSTALLDIR="%ProgramFiles%\Adobe" --INSTALLLANGUAGE=en_US'
    [string]$UninstallScript = 'msiexec /x <PackageName>.msi /qn /norestart'
    [ValidateSet('Managed', 'SelfService', 'ManagedSDL')]
    [string]$InstallType = 'Managed'
    [bool]$SSO = $true
    
    # Configuration and paths
    [hashtable]$Preferences
    [string]$TempPath
    [string]$BuildSpacePath

    # Default constructor
    AdobeApplication() { 
        $this.Init(@{}) 
    }
    
    # Constructor with properties
    AdobeApplication([hashtable]$Properties) { 
        $this.Init($Properties) 
    }

    # Shared initializer method
    [void] Init([hashtable]$Properties) {
        # Load preferences
        try {
            $this.Preferences = Get-Content "$GLOBAL:SCRIPT_ROOT\Preferences.yaml" | ConvertFrom-Yaml
            $this.TempPath = $this.Preferences.Temp
            $this.BuildSpacePath = $this.Preferences.Buildspace
        }
        catch {
            throw "Failed to load preferences: $_"
        }
        
        # Apply provided properties
        foreach ($Property in $Properties.Keys) {
            if ($this.PSObject.Properties.Name -contains $Property) {
                $this.$Property = $Properties.$Property
            }
            else {
                Write-Warning "Property '$Property' does not exist on AdobeApplication class"
            }
        }
    }
    
    <#
    .SYNOPSIS
    Updates the application version information and prepares package details.
    
    .DESCRIPTION
    Fetches the latest version information from the specified URL using regex matching,
    then updates the PackageName and UninstallScript with the current version.
    #>
    [void] Update() {
        if (-not $this.URL -or -not $this.VersionMatchStringRegex) {
            throw "URL and VersionMatchStringRegex must be set before calling Update()"
        }
        
        # Fetch version information
        try {
            $curlPath = Join-Path $this.Preferences.Tools "curl.exe"
            if (-not (Test-Path $curlPath)) {
                throw "curl.exe not found at: $curlPath"
            }
            
            Write-Log "Fetching version information from: $($this.URL)"
            $HTML = Invoke-Expression "$curlPath -s `"$($this.URL)`""
            
            if (-not $HTML) {
                throw "No content received from URL"
            }
        }
        catch {
            throw "Cannot download version HTML for Adobe Application $($this.Name): $_"
        }
        
        # Extract version using regex
        try {
            Write-Log "Extracting version using regex: $($this.VersionMatchStringRegex)"
            if ($HTML -match $this.VersionMatchStringRegex) {
                $this.Version = $matches[0]
                Write-Log "Found version: $($this.Version)"
            }
            else {
                throw "Version regex pattern did not match any content"
            }
        }
        catch {
            throw "Cannot match regular expression $($this.VersionMatchStringRegex): $_"
        }
        
        # Update package name and scripts with version
        try {
            if ($this.PackageName -and $this.PackageName.Contains("<Version>")) {
                $this.PackageName = $this.PackageName -replace "<Version>", $this.Version
                Write-Log "Updated PackageName: $($this.PackageName)"
            }
            
            if ($this.UninstallScript -and $this.UninstallScript.Contains("<PackageName>")) {
                $this.UninstallScript = $this.UninstallScript -replace "<PackageName>", $this.PackageName
                Write-Log "Updated UninstallScript: $($this.UninstallScript)"
            }
        }
        catch {
            throw "Failed to update package name or uninstall script: $_"
        }
    }

    <#
    .SYNOPSIS
    Downloads the Adobe Creative Cloud package using automated browser interaction.
    
    .DESCRIPTION
    Uses Selenium WebDriver to automate the Adobe Admin Console for package creation and download.
    Supports different installation types: Managed, SelfService, and ManagedSDL.
    #>
    [void] Download() {
        $zipfilename = "$($this.PackageName)_en_US_WIN_64.zip"
        $tempDownloadPath = Join-Path $this.TempPath $this.ID
        $driver = $null
        
        try {
            # Ensure temp directory exists
            if (-not (Test-Path $tempDownloadPath)) {
                New-Item -Path $tempDownloadPath -ItemType Directory -Force | Out-Null
            }
            
            # Initialize browser and authenticate
            $driver = $this.InitializeBrowser($tempDownloadPath)
            $this.AuthenticateToAdobe($driver)
            
            # Navigate to packages and create package
            $this.NavigateToPackages($driver)
            $this.CreatePackage($driver)
            
            # Wait for download and process
            $this.WaitForDownload($tempDownloadPath, $zipfilename)
            $this.ExtractAndMovePackage($tempDownloadPath, $zipfilename)
            
            Write-Log "Adobe package download completed successfully"
        }
        catch {
            Write-Log "ERROR: Adobe package download failed: $_"
            throw
        }
        finally {
            if ($driver) {
                try {
                    $driver.Close()
                    $driver.Quit()
                }
                catch {
                    Write-Log "WARNING: Unable to close Selenium driver cleanly"
                }
            }
        }
    }
    
    <#
    .SYNOPSIS
    Initializes the Selenium WebDriver for browser automation.
    #>
    [object] InitializeBrowser([string]$DownloadPath) {
        $consoleURL = "https://adminconsole.adobe.com"
        
        try {
            Write-Log "Starting Firefox WebDriver..."
            $driver = Start-SeDriver -Browser "Firefox" -StartURL $consoleURL -DefaultDownloadPath $DownloadPath
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
    [void] AuthenticateToAdobe([object]$Driver) {
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
        if ($this.SSO) {
            try {
                Write-Log "Authenticating via SSO..."
                Invoke-SSOAutoSignIn -Target "adminconsole.adobe.com"
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
    [void] NavigateToPackages([object]$Driver) {
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
    Creates the Adobe package based on the specified installation type.
    #>
    [void] CreatePackage([object]$Driver) {
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

            # Configure package based on install type
            switch ($this.InstallType) {
                "Managed" { $this.ConfigureManagedPackage($Driver) }
                "SelfService" { $this.ConfigureSelfServicePackage($Driver) }
                "ManagedSDL" { $this.ConfigureManagedSDLPackage($Driver) }
                default { throw "Unsupported install type: $($this.InstallType)" }
            }
        }
        catch {
            throw "Failed to create package: $_"
        }
    }
    
    <#
    .SYNOPSIS
    Configures a managed package with specific settings.
    #>
    [void] ConfigureManagedPackage([object]$Driver) {
        try {
            Write-Log "Configuring Managed package..."
            
            # Named user licensing - click next
            $this.ClickNextButton($Driver, "Named user licensing")
            
            # Select managed package
            Write-Log "Selecting managed package option..."
            $managedInput = Get-SeElement -By XPath "//*[text()='Managed package']/../following-sibling::*/input"
            $managedInput.click()
            Start-Sleep -Seconds 2
            
            $this.ClickNextButton($Driver, "Package type selection")
            $this.SelectPlatform($Driver, "Windows (64-bit)")
            $this.ClickNextButton($Driver, "Platform selection")
            $this.AddApplicationToPackage($Driver)
            $this.ClickNextButton($Driver, "Application selection")
            $this.ClickNextButton($Driver, "Skip plugins")
            
            # Configure Adobe CC Desktop Options
            Write-Log "Configuring Adobe CC Desktop options..."
            $selfServiceToggle = Get-SeElement -By XPATH "//*[text()='Enable self-service install']/../../input"
            $selfServiceToggle.click()
            Start-Sleep -Seconds 2
            
            $this.ClickNextButton($Driver, "Desktop options")
            $this.FinalizePackage($Driver)
        }
        catch {
            throw "Failed to configure managed package: $_"
        }
    }
    
    <#
    .SYNOPSIS
    Configures a self-service package.
    #>
    [void] ConfigureSelfServicePackage([object]$Driver) {
        try {
            Write-Log "Configuring Self-Service package..."
            
            # Named user licensing - click next
            $this.ClickNextButton($Driver, "Named user licensing")
            
            # Self Service is default - click next
            $this.ClickNextButton($Driver, "Self-service default")
            
            # Fill package name first for self-service
            $this.FillPackageName($Driver)
            $this.SelectPlatform($Driver, "Windows (64-bit)")
            
            # Create package
            Write-Log "Creating Self-Service package..."
            $createButton = Get-SeElement -By XPATH "//*/button[@data-testid='cta-button']"
            $createButton.Click()
        }
        catch {
            throw "Failed to configure self-service package: $_"
        }
    }
    
    <#
    .SYNOPSIS
    Configures a managed SDL (Shared Device License) package.
    #>
    [void] ConfigureManagedSDLPackage([object]$Driver) {
        try {
            Write-Log "Configuring Managed SDL package..."
            
            # Change to shared device licensing
            Write-Log "Selecting shared device licensing..."
            $sharedDeviceInput = Get-SeElement -By Xpath "//input[@aria-label='Shared device licensing']"
            $sharedDeviceInput.Click()
            Start-Sleep -Seconds 2
            
            $this.ClickNextButton($Driver, "Shared device licensing")
            
            # Select entitlement
            Write-Log "Selecting entitlement..."
            $entitlementInput = Get-SeElement -By CSSSelector ".src2-app-features-packages-components-create-package-modal-screens-choose-entitlements-page-entitlement-card-___EntitlementCard__entitlement-card___H_zoc > div:nth-child(1) > div:nth-child(2) > label:nth-child(1) > input:nth-child(1)"
            $entitlementInput.Click()
            Start-Sleep -Seconds 2
            
            $this.ClickNextButton($Driver, "Entitlement selection")
            $this.SelectPlatform($Driver, "Windows (64-bit)")
            $this.ClickNextButton($Driver, "Platform selection")
            $this.AddApplicationToPackage($Driver)
            $this.ClickNextButton($Driver, "Application selection")
            $this.ClickNextButton($Driver, "Skip plugins")
            
            # Configure desktop options
            Write-Log "Configuring desktop options for SDL..."
            $selfServiceToggle = Get-SeElement -By XPATH "//*[text()='Enable self-service install']/../../input"
            $selfServiceToggle.click()
            Start-Sleep -Seconds 2
            
            $this.ClickNextButton($Driver, "Desktop options")
            $this.FinalizePackage($Driver)
        }
        catch {
            throw "Failed to configure managed SDL package: $_"
        }
    }
    <#
    .SYNOPSIS
    Helper method to click the Next button with logging.
    #>
    [void] ClickNextButton([object]$Driver, [string]$Context) {
        Write-Log "Clicking Next button - $Context"
        $nextButton = Get-SeElement -By CSSSelector "button.Dniwja_spectrum-Button:nth-child(3) > span:nth-child(1)"
        $nextButton.click()
        Start-Sleep -Seconds 2
    }
    
    <#
    .SYNOPSIS
    Helper method to select platform (Windows 64-bit).
    #>
    [void] SelectPlatform([object]$Driver, [string]$Platform) {
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
    Helper method to add the application to the package.
    #>
    [void] AddApplicationToPackage([object]$Driver) {
        Write-Log "Adding application '$($this.Name)' to package"
        $addButton = Get-SeElement -By XPath "//*[text()=`"$($this.Name)`"]/../../../following-sibling::*/button[@data-testid='add-product-button']"
        $addButton.click()
        Start-Sleep -Seconds 2
    }
    
    <#
    .SYNOPSIS
    Helper method to fill in the package name.
    #>
    [void] FillPackageName([object]$Driver) {
        Write-Log "Filling package name: $($this.PackageName)"
        $packageNameField = Get-SeElement -By XPATH "//*/input[@data-testid='package-name-input']"
        $packageNameField.Clear()
        Invoke-SeKeys -Element $packageNameField -Keys $this.PackageName
        Start-Sleep -Seconds 1
    }
    
    <#
    .SYNOPSIS
    Helper method to finalize package creation.
    #>
    [void] FinalizePackage([object]$Driver) {
        $this.FillPackageName($Driver)
        Write-Log "Creating package..."
        $createButton = Get-SeElement -By XPATH "//*/button[@data-testid='cta-button']"
        $createButton.click()
    }
    
    <#
    .SYNOPSIS
    Waits for the package download to complete.
    #>
    [void] WaitForDownload([string]$DownloadPath, [string]$FileName) {
        $fullPath = Join-Path $DownloadPath $FileName
        Write-Log "Waiting for download to complete: $fullPath"
        
        # Wait for download to start
        $timeout = 300  # 5 minutes
        $elapsed = 0
        while (-not (Test-Path $fullPath) -and $elapsed -lt $timeout) {
            Write-Log "Waiting for download to start... ($elapsed/$timeout seconds)"
            Start-Sleep -Seconds 10
            $elapsed += 10
        }
        
        if (-not (Test-Path $fullPath)) {
            throw "Download did not start within $timeout seconds"
        }
        
        # Wait for download to complete
        $previousSize = 0
        $stableCount = 0
        while ($stableCount -lt 3) {  # File size must be stable for 3 checks
            Start-Sleep -Seconds 30
            $currentSize = (Get-Item $fullPath -ErrorAction SilentlyContinue)?.Length ?? 0
            
            if ($currentSize -eq $previousSize -and $currentSize -gt 0) {
                $stableCount++
                Write-Log "Download appears stable (check $stableCount/3) - Size: $currentSize bytes"
            }
            else {
                $stableCount = 0
                $previousSize = $currentSize
                Write-Log "Download in progress - Size: $currentSize bytes"
            }
            
            $elapsed += 30
            if ($elapsed -gt 1800) {  # 30 minutes max
                throw "Download timeout after 30 minutes"
            }
        }
        
        Write-Log "Download completed successfully!"
    }
    
    <#
    .SYNOPSIS
    Extracts the downloaded package and moves it to the build space.
    #>
    [void] ExtractAndMovePackage([string]$TempPath, [string]$ZipFileName) {
        $zipPath = Join-Path $TempPath $ZipFileName
        $extractPath = $TempPath
        $buildDestination = Join-Path $this.BuildSpacePath "$($this.ID)\$($this.Version)"
        
        try {
            Write-Log "Extracting package: $zipPath"
            Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
            
            $sourcePath = Join-Path $TempPath "$($this.PackageName)\Build\*"
            if (-not (Test-Path (Join-Path $TempPath "$($this.PackageName)\Build"))) {
                throw "Expected build directory not found after extraction"
            }
            
            # Ensure destination directory exists
            if (-not (Test-Path $buildDestination)) {
                New-Item -Path $buildDestination -ItemType Directory -Force | Out-Null
            }
            
            Write-Log "Moving package contents to build space: $buildDestination"
            Move-Item -Path $sourcePath -Destination $buildDestination -Force
            
            Write-Log "Package extraction and move completed successfully"
        }
        catch {
            throw "Failed to extract and move package: $_"
        }
    }
}