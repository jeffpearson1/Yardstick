<div align="center">
   <picture>
   <source media="(prefers-color-scheme: dark)" srcset="Branding/yardstick_logo_white_text_transparent_bg.png">
   <source media="(prefers-color-scheme: light)" srcset="Branding/yardstick_logo_transparent.png">
   <img alt="Yardstick Logo" src="Branding/yardstick_logo_transparent.png" style="width: 350px;">
   </picture>
</div>

# Yardstick

A simple, robust, easy-to-use and configure application autopackager for Microsoft Intune. 

## Description

Yardstick strives to measure up to (and beyond) the  autopackagers that already exist for MDMs without all the unnecessary complexity. 
If you don't enjoy editing XML, babysitting scripts, long days of installing applications by hand, etc. and you use Microsoft Intune - Yardstick may be for you!
Yardstick includes a variety of recipes to use directly or adapt. Its Microsoft Graph and Intune Win32 app integration is maintained in this repository; please open a feature request when a workflow is missing.

## Getting Started

You need Windows 10/11 (or Windows Server 2019+), [PowerShell 7](https://aka.ms/powershell) and an Intune tenant. The first setup also needs an account that can create Entra app registrations and grant admin consent; if you don't have one, see [Installation.md](Docs/Installation.md#creating-the-app-registration-yourself).

1. **Get Yardstick.** Clone the repository, or download the zip and extract it, into a folder with a few GB free for staging installers:
   ```powershell
   git clone https://github.com/jeffpearson1/Yardstick.git
   cd Yardstick
   ```
2. **Run the installer** from PowerShell 7, in a normal window (not **Run as administrator**), as the account that will run Yardstick:
   ```powershell
   .\Install-Yardstick.ps1
   ```
   It installs the `powershell-yaml` module, creates `Local\preferences.yaml` and the working folders, creates the Entra app registration with the Microsoft Graph permissions Yardstick needs (storing its secret in Windows Credential Manager), and optionally sets up email reports and a scheduled task (every day, on weekdays, or on days you choose). It finishes by running the health check.
3. **Package your first app:**
   ```powershell
   .\Yardstick.ps1 -AppId 7zip -NoEmail
   ```

Run `.\Test-YardstickSetup.ps1` at any time to check the installation. Re-run `.\Install-Yardstick.ps1` (or one step of it, such as `-Step AppRegistration` to renew the client secret) to change a setting; it skips whatever is already done.

The folder is also a PowerShell module. If you prefer commands to scripts, `Import-Module .\Yardstick.psd1` gives you `Invoke-Yardstick`, `Install-Yardstick` and `Test-YardstickSetup`, which take the same parameters as the scripts:
```powershell
Import-Module .\Yardstick.psd1
Invoke-Yardstick -AppId 7zip -NoEmail
```
A failed run then raises an error instead of exiting, and a stuck recipe is logged rather than closing your PowerShell window.

[Installation.md](Docs/Installation.md) covers each step in detail: the Graph permissions and why they are needed, creating the app registration by hand, scheduled-task accounts, email, upgrading and troubleshooting.

### Dependencies

The only required PowerShell Gallery module is [Powershell-Yaml](https://github.com/cloudbase/powershell-yaml), which the installer installs. Selenium-Powershell is optional, for recipes that drive a browser; it needs 4.0.0-preview3 or later, a prerelease, which the installer offers to install. TUN.CredentialManager is optional too, for recipes of your own that read a stored Windows credential with `Get-StoredCredential`; none of the shipped recipes need it. Running the unit tests needs Pester 5 or later (Windows only includes 3.4), which the installer also offers. Yardstick calls Microsoft Graph directly and has no IntuneWin32App or MSAL.PS dependency.

The official Microsoft Win32 Content Prep Tool is downloaded on first use from a commit-pinned Microsoft release and verified with SHA-256 before execution.

### Configuring preferences

Settings live in `Local\preferences.yaml`, which is excluded from Git. Every setting has a built-in default, so the file only needs what you want to change; `Config\Preferences.example.yaml` lists them all with explanations. The Intune tenant ID, client ID and secret are **not** stored in this file - they are kept in Windows Credential Manager for the account that runs Yardstick.

* Email reports go out through Outlook or SMTP. See [EmailNotifications.md](Docs/EmailNotifications.md).
* `SoftwareDropbox` and `SoftwareArchive` support recipes for vendors whose installer cannot be downloaded automatically. See [SoftwareDropbox.md](Docs/SoftwareDropbox.md).
* Setting the optional `Backup` folder keeps a copy of every `.intunewin` file Yardstick uploads, so a bad release can be traced back to the exact package that shipped.
* Site-specific recipe values, such as scope tags, go in `Local\recipeOverrides.yaml` rather than in the recipes, so the recipes can be shared and updated with `git pull`. See [Site-Specific Overrides](Docs/WritingRecipes.md#site-specific-overrides).
* Several machines can share one icon folder and one overrides file: with `syncSharedContent: true`, every run and every recipe harness script syncs them in both directions, checking each file so a bad copy never interrupts a run. See [Shared overrides](Docs/WritingRecipes.md#shared-overrides).
* A development-only computer (`.\Install-Yardstick.ps1 -DevelopmentOnly`) writes and tests recipes with the [recipe harness](Tests/RecipeHarness/README.md) and never connects to Intune. See [Development-only instances](Docs/Installation.md#development-only-instances).

### Icons

Icons are not included for licensing reasons. Put the icons your recipes need in the `Icons` folder (.jpg or .png, up to Intune's 512x512 and 750KB), or run `.\Get-YardstickIcon.ps1` to search for, tidy up and save them. The health check lists recipes whose icon is missing.

### Recipe secrets

Recipes that call a vendor API with their own keys read them with `Get-Secrets -VaultName <name>`. Store each vault in Credential Manager with `.\Set-YardstickCredential.ps1 -Secret <name>`, as the account that runs Yardstick. See [Recipe secrets](Docs/Installation.md#recipe-secrets).

### GitHub API token

Many recipes check GitHub for their latest release. GitHub allows 60 anonymous calls an hour, which one full run nearly uses up, so store a read-only token with `.\Set-YardstickCredential.ps1 -GitHub` to raise that to 5,000. See [GitHub API token](Docs/Installation.md#github-api-token).

### Recipe Tips and Tricks
* Use the defaults (configurable in preferences.yaml) for as much stuff as you can. All the available default settings are in the example preferences.yaml file.
* The installScript, uninstallScript, detectionScript and registryDetectionKey have an extra function - if you use ```<version>```, ```<filename>```, or ```<productcode>``` in them it will be replaced with the appropriate value after all the parameters and defaults are imported, processed, and the preDownloadScript has run.
* You can use recipes for locally hosted files as well - even if they are in a file share. Just define a custom downloadScript to make sure that file retrieval is handled correctly.
* Yardstick is compatible with Selenium - the fileZilla.yaml recipe is a basic example of what can be done with this. Make sure that Selenium and any drivers/browsers you need are installed first.


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

Versions can be locked to a specific value, or set of values, by adding ```versionLock:``` to a recipe, or to the recipe's entry in `Local\recipeOverrides.yaml` to lock it for your site only. There is no preference for it. Valid characters are numbers (0-9), decimals (.) and x. When the value is processed, x will be replaced with numbers of any length. 
Example: ```versionLock: 19.42.2.x``` will match versions ```19.42.2.24335``` and ```19.42.2.1``` but not ```19.42.3.442```.

### Date Offsets

Configurable in both the default and application-specific preferences files, ```deadlineDateOffset``` and ```availableDateOffset``` (as well as their defaults) will clone deployment times of application assignments and offset them forward the configured number of days.

### Supersedence & Auto-Update

Yardstick uses Intune's native update path instead of a custom remediation. After each upload:

* The new version declares **supersedence** over every older version Yardstick keeps (and over the DETECT anchor, when one exists). ```uninstallPreviousVersion``` selects between Intune's `Update` (in-place upgrade, default) and `Replace` (uninstall first) behavior.
* **Auto-update** (```autoUpdateSupersededAppsState```) is turned on for the new version's *available* assignments, so Intune pulls devices running a superseded version forward without any user action. Intune only supports auto-update for available assignments, so required-only recipes are unaffected. Specific device or user groups can opt out via ```defaultGroupsSkipAutoUpdates``` in ```preferences.yaml``` or ```groupSkipAutoUpdates``` in a recipe; assignments targeting those groups are held at ```notConfigured```.
* **Required** assignments are moved onto the new version; **available** assignments are copied and left in place, because removing an available assignment destroys the on-device component Intune uses to auto-update. If the older version will not be superseded (```supersedence: false```), available assignments are moved instead, as in earlier releases.

Defaults live in ```Local\preferences.yaml``` (```defaultSupersedence```, ```defaultUninstallPreviousVersion```, ```defaultAutoUpdate```, ```useDetectAnchor```) and can be overridden per recipe. See [WritingRecipes.md](Docs/WritingRecipes.md) for details, including the DETECT anchor that keeps long-abandoned installs in scope.

Existing tenants can be migrated onto this model in one pass:

```powershell
.\Migrate-ToSupersedenceModel.ps1 -WhatIf
```

#### Native Graph integration

`Modules/YardstickGraph.psm1` handles OAuth client credentials, token refresh, transient retries, and OData paging. `Modules/YardstickIntune.psm1` owns Win32 app packaging, upload, CRUD, assignments, dependencies, and supersedence. Assignment collections are normalized correctly for zero, one, or many results.

#### Known issue: supersedence direction is not carried by `sourceId`

The beta `mobileApps/{id}/relationships` collection is reported from the perspective of the app being queried: `sourceId` is the queried app on **both** forward and reverse links, and `targetId` is always the app at the other end. Direction lives in `targetType` — `child` means the queried app supersedes the target, `parent` means the target supersedes the queried app. Classifying by `sourceId` marks every link as forward, so a superseded app is never detached from its parent and its deletion fails with `supersedes <parent id>`. `Get-YardstickSupersedenceRelationship` keys off `targetType`.

#### Known issue: un-targeting a group requires editing the anchor too

The DETECT anchor participates in assignment migration like any other older version — its **required** assignments are moved onto the new version, its **available** assignments are copied and left in place. Unlike an (N-x) version, though, the anchor is never pruned, so its available assignments are re-copied onto every subsequent release indefinitely.

The practical consequence: removing an available assignment from the current version alone does **not** un-target that group, because the next Yardstick run re-creates it from the anchor. To drop a group permanently, remove the assignment from the DETECT anchor as well.

#### Other Parameters

* ```-Force``` will overwrite the latest version of any targeted applications if they are the same as the new version, and run normally if a new version is available.
* ```-Repair``` runs the full maintenance sweep — (N-x) renaming, assignment migration, retention pruning, supersedence and auto-update — on applications that have **no update available**, which a normal run skips entirely. See [Maintenance sweeps](#maintenance-sweeps) below.
* ```-NoDelete``` will stop the script from automatically deleting old versions when it is done. 
* ```-NoInteractive``` will skip any recipes that are located in the "Interactive" folder in Recipes.
* ```-Group``` will parse RecipeGroups.yaml and run any applications corresponding to the group provided

#### Keeping a recipe group in sync

`Sync-RecipeGroup.ps1` rewrites one group in `RecipeGroups.yaml` so it lists exactly the recipes in the matching folder. It defaults to the Development group and `Recipes\Development`, ignores recipes in a `Disabled` folder, and leaves every other group untouched.

```powershell
.\Sync-RecipeGroup.ps1 -WhatIf   # show what would be added and removed
.\Sync-RecipeGroup.ps1
.\Sync-RecipeGroup.ps1 -GroupName TestApps -RecipeFolder .\Recipes\Testing
```

#### Resending a run's email report

If a run's email was skipped or failed, `Send-YardstickLogReport.ps1` rebuilds the report from `YLog.log` and sends it. See [Docs/EmailNotifications.md](Docs/EmailNotifications.md#resending-a-report-from-the-log).

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

