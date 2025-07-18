# Yardstick

A simple, robust, easy-to-use and configure application autopackager for Microsoft Intune. 

## Description

Yardstick strives to measure up to (and beyond) the  autopackagers that already exist for MDMs without all the unnecessary complexity. 
If you don't enjoy editing XML, babysitting scripts, long days of installing applications by hand, etc. and you use Microsoft Intune - Yardstick may be for you!
We have included a variety of recipes for you to either use directly or modify, and if you find that there is functionality missing from what you would expect (especially if it is something already present in IntuneWin32App) please submit a feature request so we can look into getting it added.

## Getting Started

### Dependencies

Yardstick depends on (and we are extremely grateful for) a handful of PowerShell Modules:
* [Powershell-Yaml](https://github.com/cloudbase/powershell-yaml)
* [Selenium-Powershell](https://github.com/adamdriscoll/selenium-powershell)
* [IntuneWin32App](https://github.com/MSEndpointMgr/IntuneWin32App)
* [PowerShell_Credential_Manager](https://github.com/echalone/PowerShell_Credential_Manager)


### Installing

* Before starting, install all necessary PowerShell Modules
    * Install the latest version of our modified IntuneWin32App module from [this repo](https://github.com/jeffpearson1/IntuneWin32App)
    * The remaining modules can be installed from the PowerShell Gallery:

```powershell
Install-Module -Name Powershell-Yaml, TUN.CredentialManager
Install-Module -Name Selenium -AllowPrerelease
```

* Download or clone the repository into a folder where it can live. A sufficient amount of disk space should be available in this directory for the staging of applications. 
* Some recipes may require you to have Mozilla Firefox installed and configured to work with Selenium.


### Configuring Preferences.yaml

Most of this file should be fairly self-explanatory. The TenantID, ClientID, and ClientSecrets are required for Yardstick to operate properly. Defaults do not necessarily have to be set, however recipes that don't contain all the values normally set by the defaults may fail to run correctly.


### Populate the icon cache

Icons are not included for licensing reasons. Populate the icon cache folder with the icons needed for any application recipes you will be running. Formats can be .jpg or .png, max size is the same as Intune - 512x512 and 750KB. Be sure to double check filename extensions in recipes you are using.


### Set credentials in Windows Credential Manager

Create a new Windows credential - ```yourdomainnamehere``` that contains the username and password that will be used to sign into the Intune Graph API, along with any other credential objects that may be required by recipes you are running.


### Recipe Tips and Tricks
* Use the defaults (configurable in preferences.yaml) for as much stuff as you can. All the available default settings are in the example preferences.yaml file.
* The installScript, uninstallScript, detectionScript and registryDetectionKey have an extra function - if you use ```<version>```, ```<filename>```, or ```<productcode>``` in them it will be replaced with the appropriate value after all the parameters and defaults are imported, processed, and the preDownloadScript has run.
* You can use recipes for locally hosted files as well - even if they are in a file share. Just define a custom downloadScript to make sure that file retrieval is handled correctly.
* Yardstick is compatible with Selenium - the filezilla.yaml recipe is a basic example of what can be done with this. Make sure that Selenium and any drivers/browsers you need are installed first.


### Running Yardstick

* You can either run all applications at once
```powershell
.\Yardstick.ps1 -All
```
* or you can run individual applications with their "id" (the name of the recipe file without the extension)
```powershell
.\Yardstick.ps1 -AppId googlechrome
```

### Version Locking

Versions can be locked to a specific value, or set of values by adding the ```versionLock:``` parameter to any preferences file. Valid characters are numbers (0-9), decimals (.) and x. When the value is processed, x will be replaced with numbers of any length. 
Example: ```versionLock: 19.42.2.x``` will match versions ```19.42.2.24335``` and ```19.42.2.1``` but not ```19.42.3.442```.

### Date Offsets

Configurable in both the default and application-specific preferences files, ```deadlineDateOffset``` and ```availableDateOffset``` (as well as their defaults) will clone deployment times of application assignments and offset them forward the configured number of days.

#### Other Parameters

* ```-Force``` will overwrite the latest version of any targeted applications if they are the same as the new version, and run normally if a new version is available.
* ```-Repair``` will fix any name discrepancies of (N-X) for any target applications (i.e. if multiple applications are named N-1 - although this is normally fixed after an update anyway).
* ```-NoDelete``` will stop the script from automatically deleting old versions when it is done. 
* ```-NoInteractive``` will skip any recipes that are located in the "Interactive" folder in Recipes.
* ```-Group``` will parse RecipeGroups.yaml and run any applications corresponding to the group provided


## License

This project is licensed under the MIT License - see the LICENSE.md file for details


## Inspiration

* [CMPackager](https://github.com/asjimene/CMPackager)

