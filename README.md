# Yardstick

A simple, robust, easy-to-use and configure application autopackager for Microsoft Intune. 

## Description

Yardstick strives to measure up to (and beyond) the plethora of autopackagers that already exist for MDMs without all the unnecessary complexity. 
If you don't enjoy editing XML, babysitting scripts, long days of installing applications by hand, etc. and you use Microsoft Intune - Yardstick may be for you!
We have included a variety of recipes for you to either use directly or modify, and if you find that there is functionality missing from what you would expect (especially if it is something already present in IntuneWin32App) please submit a feature request so we can look into getting it added.

## Getting Started

### Dependencies

Yardstick depends on (and is extremely grateful for) a handful of PowerShell Modules:
* [Powershell-Yaml](https://github.com/cloudbase/powershell-yaml)
* [Selenium-Powershell](https://github.com/adamdriscoll/selenium-powershell)
* [IntuneWin32App](https://github.com/MSEndpointMgr/IntuneWin32App)

### Installing

* Before starting, install all necessary PowerShell 
```powershell
Install-Module -Name Powershell-Yaml, IntuneWin32App, Selenium-PowerShell
```
* Download or clone the repository into a folder where it can live. A sufficient amount of disk space should be available in this directory for the staging of applications. Currently Yardstick keeps old application files in a directory (i.e. BuildSpace\Old) which will need to be cleaned up every once in a while.

### Configuring Preferences.yaml

Most of this file should be fairly self-explanatory. The TenantID, ClientID, and ClientSecrets are required for Yardstick to operate properly. Defaults do not necessarily have to be set, however recipes that don't contain all the values normally set by the defaults may fail to run correctly.

### Executing program

* You can either run all applications at once
```powershell
.\Yardstick.ps1 -All
```
* or you can run individual applications with their "id" (the name of the recipe file without the extension)
```powershell
.\Yardstick.ps1 -AppId googlechrome
```


## Version History

* 0.1
    * Initial Release

## License

This project is licensed under the MIT License - see the LICENSE.md file for details

## Inspiration

* [CMPackager](https://github.com/asjimene/CMPackager)

