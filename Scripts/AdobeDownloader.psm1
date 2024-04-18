function Start-AdobeDownload {
    param(
        [string[]] $SearchNames
    )
    try {
        $zipfilename = "$($Script:PackageName)_en_US_WIN_64.zip"
        $ConsoleURL = "https://adminconsole.adobe.com"
        Write-Log "Tempdir: $Script:tempDir"
        Write-Log "ID: $Script:id and $id"
        $Script:Driver = Start-SeDriver -Browser "Firefox" -StartURL $ConsoleURL -DefaultDownloadPath "$Script:tempDir\$Script:id" -Arguments @('--headless', '--window-size=1920,1080')
        Start-Sleep -Seconds 15
        # Load the credentials from the password manager
        $Credentials = Get-StoredCredential -Target "adobe"
        # Enter the username into the sign-in dialog
        $Element = Get-SeElement -By ID "EmailPage-EmailField"
        Invoke-SeKeys -Element $Element -Keys $Credentials.Username
        $Button = Get-SeElement -By CSSSelector ".spectrum-Button"
        $Button.click()
        Start-Sleep -Seconds 15
        # Sign in via SSO
        Invoke-SSOAutoSignIn -Target "adminconsole.adobe.com"
        # Go to the packages tab
        while ( -not (Get-SeElement -By LinkText "Overview" -ErrorAction SilentlyContinue)) {
            Write-Log "Waiting for page load..."
            Start-Sleep 3
        }
        Write-Log "Selecting Packages tab..."
        (Get-SeElement -By LinkText "Packages").click()
        
        # Click the create a package button
        While ( -not (Get-SeElement -By CSSSelector "button.Dniwja_spectrum-Button:nth-child(2)" -ErrorAction SilentlyContinue)) {
            Write-Log "Waiting for page load..."
            Start-Sleep 3
        }
        Write-Log "Clicking 'create a package'"
        (Get-SeElement -By CSSSelector "button.Dniwja_spectrum-Button:nth-child(2)").click()
        Start-Sleep -Seconds 5
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
        # Add application(s) to list
        Write-Log "Add app to list"
        foreach ($Name in $SearchNames) {
            (Get-SeElement -By XPath "//*[text()=`"$Name`"]/../../../following-sibling::*/button[@data-testid='add-product-button']").click()
            Start-Sleep -Seconds 2
        }
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
        Write-Log "Fill out package name with $Script:PackageName"
        $PackageName_Field = Get-SeElement -By XPATH "//*/input[@data-testid='package-name-input']"
        Invoke-SeKeys -Element $PackageName_Field -Keys $Script:PackageName
        # Click Create package
        (Get-SeElement -By XPATH "//*/button[@data-testid='cta-button']").click()
        Write-Log "Package is building and will download automatically..."
        Write-Log "Filename: $zipfilename"
        while (!(Test-Path $Script:tempDir\$Script:id\$zipfilename)) {
        Write-Log "Waiting for download to start..."
        Start-Sleep -Seconds 5
        }
        while ((Get-Item $Script:tempDir\$Script:id\$zipfilename).Length -le 0) {
        Write-Log "Waiting for download to finish..."
        Start-Sleep -Seconds 30
        }
        Write-Log "Download complete!"
        # Close Firefox
        $Driver.Close()
        # Extract the zip file
        Write-Log "Expanding downloaded archive..."
        Expand-Archive $Script:tempDir\$Script:id\$zipfilename -DestinationPath $Script:tempDir\$Script:id
        # Move into buildspace
        Write-Log "Moving contents into buildspace... ($Script:tempDir\$Script:id\$Script:packageName\Build\* to $Script:buildspace\$Script:id\$Script:version)"
        Move-Item $Script:tempDir\$Script:id\$Script:packageName\Build\* $Script:buildspace\$Script:id\$Script:version -Force
    }
    catch {
        Write-Error "ERROR: There was an issue running Start-AdobeDownload for application $SearchName"
        return 1
    }
}