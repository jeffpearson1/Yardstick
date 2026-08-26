<p align="center">
  <img src="Branding/yardstick_logo_transparent.png" alt="Yardstick logo" width="300">
</p>

# Yardstick

A simple, robust, easy-to-use and configure application autopackager for Microsoft Intune. 

## Description

Yardstick strives to measure up to (and beyond) the  autopackagers that already exist for MDMs without all the unnecessary complexity. 
If you don't enjoy editing XML, babysitting scripts, long days of installing applications by hand, etc. and you use Microsoft Intune - Yardstick may be for you!
Yardstick includes a variety of recipes to use directly or adapt. Its Microsoft Graph and Intune Win32 app integration is maintained in this repository; please open a feature request when a workflow is missing.

## Getting Started

### Dependencies

The only required PowerShell Gallery module is [Powershell-Yaml](https://github.com/cloudbase/powershell-yaml). Selenium-Powershell and TUN.CredentialManager remain optional for recipes that use browser automation or stored vendor credentials. Yardstick calls Microsoft Graph directly and has no IntuneWin32App or MSAL.PS dependency.

The official Microsoft Win32 Content Prep Tool is downloaded on first use from a commit-pinned Microsoft release and verified with SHA-256 before execution.


### Installing

```powershell
Install-Module -Name Powershell-Yaml
# Optional, for recipes that use these integrations:
Install-Module -Name TUN.CredentialManager
Install-Module -Name Selenium -AllowPrerelease
```

* Download or clone the repository into a folder where it can live. A sufficient amount of disk space should be available in this directory for the staging of applications. 
* Some recipes may require you to have Mozilla Firefox installed and configured to work with Selenium.


### Configuring Preferences.yaml

Most of this file should be fairly self-explanatory. Defaults do not necessarily have to be set, however recipes that don't contain all the values normally set by the defaults may fail to run correctly.

The TenantID, ClientID and ClientSecret are **not** stored in this file - see [Set the Intune credentials](#set-the-intune-credentials) below. `credentialTarget`, `credentialExpirationWarningDays`, `credentialExpirationEmailIntervalHours` and `adminEmailRecipient` tune where those credentials are stored and who is told when they are about to expire.

Setting the optional `Backup` folder makes Yardstick keep a copy of every `.intunewin` file it uploads, so a bad release can be traced back to the exact package that shipped. Each recipe gets its own subfolder, and each file is named with the version and the time Intune finished publishing it. The newest three are retained (`backupVersionsToKeep`). The copy runs on a background thread while Yardstick carries on with supersedence and assignments, and a backup failure is reported without failing the application update. Leave `Backup` blank to turn this off.


### Populate the icon cache

Icons are not included for licensing reasons. Populate the icon cache folder with the icons needed for any application recipes you will be running. Formats can be .jpg or .png, max size is the same as Intune - 512x512 and 750KB. Be sure to double check filename extensions in recipes you are using.


### Set the Intune credentials

The app registration credentials Yardstick uses to reach Microsoft Graph are held in
Windows Credential Manager, encrypted for the account that runs Yardstick, rather than
in plaintext in `preferences.yaml`. Store them once per machine:

```powershell
.\Set-YardstickCredential.ps1
```

You are prompted for the Tenant ID, Client ID and Client Secret (typed as a masked
SecureString), plus an optional secret expiration date. The script then authenticates
to Graph to confirm the values work before you rely on them.

Other useful invocations:

```powershell
.\Set-YardstickCredential.ps1 -Show               # tenant, client and expiry (secret masked)
.\Set-YardstickCredential.ps1 -RefreshExpiration  # re-read the expiry date from Graph
.\Set-YardstickCredential.ps1 -Remove             # delete the stored credential
```

If `preferences.yaml` still carries the legacy `TenantID` / `ClientId` / `ClientSecret`
keys, the next Yardstick run migrates them into Credential Manager automatically and
logs a reminder to delete them from the file. Running with `-NoInteractive` never
prompts: an unattended run with no stored credential fails immediately with an
instruction to run `Set-YardstickCredential.ps1`.

#### Client secret expiration

At the start of every run Yardstick checks how much life the client secret has left.
When it is inside `credentialExpirationWarningDays` (default 30), a warning is written
to the console and log, and an email goes to `adminEmailRecipient` (falling back to
`emailRecipient`). Repeat emails are throttled to one per
`credentialExpirationEmailIntervalHours` (default 24) so a nightly schedule does not
spam the mailbox.

The expiration date is discovered from Graph when the app registration holds
`Application.Read.All` - Yardstick matches the secret to its `passwordCredential` by
hint. Without that permission, supply the date yourself when running
`Set-YardstickCredential.ps1`.


### Set recipe credentials in Windows Credential Manager

Some recipes sign in to vendor portals. Create a Windows credential -
```yourdomainnamehere``` - that contains the username and password used for those
sign-ins, along with any other credential objects required by the recipes you run.


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

### Supersedence & Auto-Update

Yardstick uses Intune's native update path instead of a custom remediation. After each upload:

* The new version declares **supersedence** over every older version Yardstick keeps (and over the `{DETECT}` anchor, when one exists). ```uninstallPreviousVersion``` selects between Intune's `Update` (in-place upgrade, default) and `Replace` (uninstall first) behavior.
* **Auto-update** (```autoUpdateSupersededAppsState```) is turned on for the new version's *available* assignments, so Intune pulls devices running a superseded version forward without any user action. Intune only supports auto-update for available assignments, so required-only recipes are unaffected. Specific device or user groups can opt out via ```defaultGroupsSkipAutoUpdates``` in ```preferences.yaml``` or ```groupSkipAutoUpdates``` in a recipe; assignments targeting those groups are held at ```notConfigured```.
* **Required** assignments are moved onto the new version; **available** assignments are copied and left in place, because removing an available assignment destroys the on-device component Intune uses to auto-update. If the older version will not be superseded (```supersedence: false```), available assignments are moved instead, as in earlier releases.

Defaults live in ```preferences.yaml``` (```defaultSupersedence```, ```defaultUninstallPreviousVersion```, ```defaultAutoUpdate```, ```useDetectAnchor```) and can be overridden per recipe. See [WritingRecipes.md](WritingRecipes.md) for details, including the `{DETECT}` anchor that keeps long-abandoned installs in scope.

Existing tenants can be migrated onto this model in one pass:

```powershell
.\Migrate-ToSupersedenceModel.ps1 -WhatIf
```

#### Native Graph integration

`Modules/YardstickGraph.psm1` handles OAuth client credentials, token refresh, transient retries, and OData paging. `Modules/YardstickIntune.psm1` owns Win32 app packaging, upload, CRUD, assignments, dependencies, and supersedence. Assignment collections are normalized correctly for zero, one, or many results.

#### Known issue: supersedence direction is not carried by `sourceId`

The beta `mobileApps/{id}/relationships` collection is reported from the perspective of the app being queried: `sourceId` is the queried app on **both** forward and reverse links, and `targetId` is always the app at the other end. Direction lives in `targetType` — `child` means the queried app supersedes the target, `parent` means the target supersedes the queried app. Classifying by `sourceId` marks every link as forward, so a superseded app is never detached from its parent and its deletion fails with `supersedes <parent id>`. `Get-YardstickSupersedenceRelationship` keys off `targetType`.

#### Known issue: un-targeting a group requires editing the anchor too

The `{DETECT}` anchor participates in assignment migration like any other older version — its **required** assignments are moved onto the new version, its **available** assignments are copied and left in place. Unlike an (N-x) version, though, the anchor is never pruned, so its available assignments are re-copied onto every subsequent release indefinitely.

The practical consequence: removing an available assignment from the current version alone does **not** un-target that group, because the next Yardstick run re-creates it from the anchor. To drop a group permanently, remove the assignment from the `{DETECT}` anchor as well.

#### Other Parameters

* ```-Force``` will overwrite the latest version of any targeted applications if they are the same as the new version, and run normally if a new version is available.
* ```-Repair``` runs the full maintenance sweep — (N-x) renaming, assignment migration, retention pruning, supersedence and auto-update — on applications that have **no update available**, which a normal run skips entirely. See [Maintenance sweeps](#maintenance-sweeps) below.
* ```-NoDelete``` will stop the script from automatically deleting old versions when it is done. 
* ```-NoInteractive``` will skip any recipes that are located in the "Interactive" folder in Recipes.
* ```-Group``` will parse RecipeGroups.yaml and run any applications corresponding to the group provided

#### Maintenance sweeps

Retention, naming and supersedence are desired state, but they used to be reconciled only on the code path that publishes a new version. A run where the vendor shipped nothing simply stopped at the version check. Anything that failed to prune — blocked by a dependent app, refused by Intune, or stranded by a bug — therefore got exactly one retry: the next release. For an application that updates rarely that can be never, which is how tenants accumulate `(N-2)`/`(N-3)` versions and supersedence graphs that hit Intune's 9-target limit.

`-Repair` now runs that same reconciliation against the newest version already in Intune:

```powershell
.\Yardstick.ps1 -All -Repair            # update what has a new version, reconcile everything else
.\Yardstick.ps1 -All -Repair -NoDelete  # reconcile everything except deletions
```

Every step is idempotent — renames are no-ops when the name is already right, assignment migration re-reads both ends, and supersedence is replaced as a whole set — so a sweep over a clean application changes nothing and is not reported. Applications where the sweep pruned, renamed or pinned an anchor appear in the email report with the action **Maintained**.

## License

This project is licensed under the MIT License - see the LICENSE.md file for details


## Inspiration

* [CMPackager](https://github.com/asjimene/CMPackager)

