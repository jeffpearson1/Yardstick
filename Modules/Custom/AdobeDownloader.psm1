class AdobeApplication {
    [string]$Name
    [string]$ID
    [string]$URL
    [string]$VersionMatchStringRegex
    [string]$VersionYear
    [string]$VersionYearLocation
    [string]$Version
    [string]$VersionLocation
    [string]$FileDetectionPath
    [string]$PackageName
    [string]$InstallScript = 'setup.exe --silent --ADOBEINSTALLDIR="%ProgramFiles%\Adobe" --INSTALLLANGUAGE=en_US'
    [string]$UninstallScript = 'msiexec /x <PackageName>.msi /qn /norestart'
    [string]$InstallType = 'Managed'
    [bool]$SSO=$true
    [array]$prefs = (Get-Content G:\Intune\Yardstick\Preferences.yaml | ConvertFrom-Yaml)

    AdobeApplication() { 
        $this.Init(@{}) 
    }
    AdobeApplication([hashtable]$Properties) { 
        $this.Init($Properties) 
    }

    # Shared initializer method
    [void] Init([hashtable]$Properties) {
        foreach ($Property in $Properties.Keys) {
            $this.$Property = $Properties.$Property
        }
    }
    

    [void] Update() {
        # try {
            $HTML = Invoke-Expression "$($this.prefs.Tools)\curl.exe -s $($this.url)"
            [string]($HTML -match $this.VersionMatchStringRegex) -match $this.VersionMatchStringRegex | Out-Null
            $VersionMatchString = $matches[0].trim()
            $VersionMatchArray = $VersionMatchString -split "\>|\ "
            if ($this.VersionLocation)
            {
                $this.Version = $VersionMatchArray[$this.VersionLocation]
            }
            if ($this.VersionYearLocation) {
                $this.VersionYear = $VersionMatchArray[$this.VersionYearLocation]
            }
            $this.PackageName = $this.PackageName -replace "<Version>", $this.Version
            $this.UninstallScript = $this.uninstallScript -replace "<PackageName>", $this.PackageName
            $this.FileDetectionPath = $this.FileDetectionPath -replace "<VersionYear>", $this.VersionYear
        # }
        # catch {
        #     Write-Error "Issue calling update for $($this.Name)"
        # }
    }

    [void] Download() {
        $zipfilename = "$($this.PackageName)_en_US_WIN_64.zip"
        $ConsoleURL = "https://adminconsole.adobe.com"
        try {
            $Driver = Start-SeDriver -Browser "Firefox" -StartURL $ConsoleURL -DefaultDownloadPath "$($this.prefs.TEMP)\$($this.id)" #-Arguments @('--headless', '--window-size=1920,1080')
            Start-Sleep -Seconds 15
        }
        catch {
            Write-Log "ERROR: Cannot start Firefox with Selenium!"
            Write-Error "ERROR: Cannot start Firefox with Selenium!"
            return
        }
        
        try {
            # Load the credentials from the password manager
            $Credentials = Get-StoredCredential -Target "adobe"
        }
        catch {
            Write-Log "ERROR: Cannot retrieve stored credentials with target 'adobe'!"
            Write-Error "ERROR: Cannot retrieve stored credentials with target 'adobe'!"
            return
        }

        try {
            # Enter the username into the sign-in dialog
            $Element = Get-SeElement -By ID "EmailPage-EmailField"
            Invoke-SeKeys -Element $Element -Keys $Credentials.Username
            $Button = Get-SeElement -By CSSSelector ".spectrum-Button"
            $Button.click()
            Start-Sleep -Seconds 15
        }
        catch {
            Write-Log "ERROR: Cannot pass Adobe username only sign in dialog!"
            Write-Error "ERROR: Cannot pass Adobe username only sign in dialog!"
            return
        }
        

        if ($this.SSO) {
            try {
                # Sign in via SSO
                Invoke-SSOAutoSignIn -Target "adminconsole.adobe.com" -Driver $Driver
            }
            catch {
                Write-Log "ERROR: Cannot run Invoke-SSOAutoSignIn!"
                Write-Error "ERROR: Cannot run Invoke-SSOAutoSignIn!"
                return
            }
        }
        else {
            try {
                # Enter the password
                $Element = Get-SeElement -By ID "PasswordPage-PasswordField"
                Invoke-SeKeys -Element $Element -Keys $((New-Object PSCredential 0, $Credentials.Password).GetNetworkCredential().Password)
                $Button = Get-SeElement -By CSSSelector ".spectrum-Button"
                $Button.click()
            }
            catch {
                Write-Log "ERROR: There was an issue with sign in"
                Write-Error "ERROR: There was an issue with sign in"
                return
            }
        }


        
        
        try {
            # Go to the packages tab
            while ( -not (Get-SeElement -By LinkText "Overview" -ErrorAction SilentlyContinue)) {
                Write-Log "Waiting for page load..."
                Start-Sleep 3
            }
            Write-Log "Selecting Packages tab..."
            (Get-SeElement -By LinkText "Packages").click()
        }
        catch {
            Write-Log "ERROR: Cannot open Packages tab in Adobe Admin Console"
            Write-Error "ERROR: Cannot open Packages tab in Adobe Admin Console"
            return
        }
        
        try {
            # Click the create a package button
            While ( -not (Get-SeElement -By CSSSelector "button.Dniwja_spectrum-Button:nth-child(2)" -ErrorAction SilentlyContinue)) {
                Write-Log "Waiting for page load..."
                Start-Sleep 3
            }
            Write-Log "Clicking 'create a package'"
            (Get-SeElement -By CSSSelector "button.Dniwja_spectrum-Button:nth-child(2)").click()
            Start-Sleep -Seconds 5
        }
        catch {
            Write-Log "ERROR: Cannot click 'create a package'"
            Write-Error "ERROR: Cannot click 'create a package'"
            return
        }

        if ($this.InstallType -eq "Managed") {
            try {
                # Named user licensing is default, so just click the next button
                Write-Log "NU: Click next"
                (Get-SeElement -By CSSSelector "div.aaz5ma_spectrum-ButtonGroup:nth-child(4) > button:nth-child(2) > span:nth-child(1)").click()
                Start-Sleep -Seconds 2
                # Pick managed package
                Write-Log "Click managed package"
                (Get-SeElement -By XPath "//*[text()='Managed package']/../following-sibling::*/input").click()
                Start-Sleep -Seconds 2
                # Click next
                Write-Log "Click Next"
                (Get-SeElement -By CSSSelector "button.Dniwja_spectrum-Button:nth-child(3) > span:nth-child(1)").click()
                Start-Sleep -Seconds 2
                # Click the OS dropdown
                Write-Log "Click OS select dropdown"
                (Get-SeElement -By XPath "//*[text()='Select platform']").click()
                Start-Sleep -Seconds 2
                # Pick Windows (64-bit)
                Write-Log "Select Windows (64-bit)"
                (Get-SeElement -By XPath "//*[text()='Windows (64-bit)']").click()
                Start-Sleep -Seconds 2
                # Click next
                Write-Log "Click Next"
                (Get-SeElement -By CSSSelector "button.Dniwja_spectrum-Button:nth-child(3) > span:nth-child(1)").click()
                Start-Sleep -Seconds 2
                # Add application to list
                Write-Log "Add app to list"
                (Get-SeElement -By XPath "//*[text()=`"$($this.Name)`"]/../../../following-sibling::*/button[@data-testid='add-product-button']").click()
                Start-Sleep -Seconds 2
                # Click next
                Write-Log "Click Next"
                (Get-SeElement -By CSSSelector "button.Dniwja_spectrum-Button:nth-child(3) > span:nth-child(1)").click()
                Start-Sleep -Seconds 2
                # Skip Plugins and click next
                Write-Log "Skip Plugins: Click Next"
                (Get-SeElement -By CSSSelector "button.Dniwja_spectrum-Button:nth-child(3) > span:nth-child(1)").click()
                Start-Sleep -Seconds 2
                # Adobe CC Desktop Options
                # Disable Auto Update for Computer Labs
                Write-Log "Disable auto update for computer labs"
                (Get-SeElement -By XPATH "//*[text()='Enable self-service install']/../../input").click()
                Start-Sleep -Seconds 2
                # Click next
                Write-Log "Click Next"
                (Get-SeElement -By CSSSelector "button.Dniwja_spectrum-Button:nth-child(3) > span:nth-child(1)").click()
                Start-Sleep -Seconds 2
                # Fill out the package name
                Write-Log "Fill out package name with $($this.PackageName)"
                $PackageName_Field = Get-SeElement -By XPATH "//*/input[@data-testid='package-name-input']"
                Invoke-SeKeys -Element $PackageName_Field -Keys $this.PackageName
                # Click Create package
                (Get-SeElement -By XPATH "//*/button[@data-testid='cta-button']").click()
            }
            catch {
                Write-Log "ERROR: There was an issue configuring new package parameters! (Managed)"
                Write-Error "ERROR: There was an issue configuring new package parameters! (Managed)"
                return
            }
        }

        if ($this.InstallType -eq "SelfService") {
            try {
                # Named user licensing is default, so just click the next button
                Write-Log "NU: Click next"
                (Get-SeElement -By CSSSelector "div.aaz5ma_spectrum-ButtonGroup:nth-child(4) > button:nth-child(2) > span:nth-child(1)").click()
                Start-Sleep -Seconds 2
                # Self Service is the default - click next
                Write-Log "Click Next"
                (Get-SeElement -By CSSSelector "button.Dniwja_spectrum-Button:nth-child(3) > span:nth-child(1)").click()
                Start-Sleep -Seconds 2
                Write-Log "Fill out package name with $($this.PackageName)"
                $PackageName_Field = Get-SeElement -By XPATH "//*/input[@data-testid='package-name-input']"
                Invoke-SeKeys -Element $PackageName_Field -Keys $this.PackageName
                Start-Sleep -Seconds 0.5
                # Click the OS dropdown
                Write-Log "Click OS select dropdown"
                (Get-SeElement -By XPath "//*[text()='Select platform']").click()
                Start-Sleep -Seconds 0.5
                # Pick Windows (64-bit)
                Write-Log "Select Windows (64-bit)"
                (Get-SeElement -By XPath "//*[text()='Windows (64-bit)']").click()
                Start-Sleep -Seconds 0.5
                # Click Create package
                (Get-SeElement -By XPATH "//*/button[@data-testid='cta-button']").click()
            }
            catch {
                Write-Log "ERROR: There was an issue configuring new package parameters! (Self Service)"
                Write-Error "ERROR: There was an issue configuring new package parameters! (Self Service)"
                return
            }
        }
        
        try {
            Write-Log "Package is building and will download automatically..."
            Write-Log "Filename: $zipfilename"
            Write-Log "Test: $($this.Prefs.Temp)\$($this.ID)\$($zipfilename)"
            while (!(Test-Path "$($this.Prefs.Temp)\$($this.ID)\$($zipfilename)")) {
            Write-Log "Waiting for download to start..."
            Start-Sleep -Seconds 5
            }
            while ((Get-Item "$($this.Prefs.Temp)\$($this.ID)\$($zipfilename)").Length -le 0) {
            Write-Log "Waiting for download to finish..."
            Start-Sleep -Seconds 30
            }
            Write-Log "Download complete!"
        }
        catch {
            Write-Log "Failed to download"
            Write-Error "Failed to download"
            return
        }
            
        
        try {
            # Close Firefox
            $Driver.Close()
        }
        catch {
            Write-Log "ERROR: Unable to close selenium driver. It may need to be quit manually."
        }
        
        try {
            # Extract the zip file
            Write-Log "Expanding downloaded archive..."
            Expand-Archive "$($This.Prefs.TEMP)\$($this.id)\$zipfilename" -DestinationPath "$($This.Prefs.TEMP)\$($this.id)"
        }
        catch {
            Write-Log "ERROR: Cannot expand downloaded archive"
            Write-Error "ERROR: Cannot expand downloaded archive"
            return
        }
        
        try {
            # Move into buildspace
            Write-Log "Moving contents into buildspace... ($($This.Prefs.TEMP)\$($this.id)\$($this.packageName)\Build\* to $($this.BUILDSPACE)\$($this.id)\$($this.version))"
            Move-Item "$($This.Prefs.TEMP)\$($this.id)\$($this.packageName)\Build\*" "$($This.Prefs.BUILDSPACE)\$($this.id)\$($this.version)" -Force
        }
        catch {
            Write-Log "ERROR: Moving contents to buildspace has failed."
            Write-Error "ERROR: Moving contents to buildspace has failed."
            return
        }
    }
}