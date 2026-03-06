# Yardstick Copilot Instructions

## Architecture & Flow
- Primary orchestration lives in [Yardstick.ps1](Yardstick.ps1); it imports IntuneWin32App, Selenium, and local support modules, builds an application list via `-ApplicationId`, `-Group`, or `-All`, and loops each recipe through download → packaging → Intune upload.
- `Connect-AutoMSIntuneGraph` in [Modules/YardstickSupport.psm1](Modules/YardstickSupport.psm1) auto-refreshes Graph tokens; call it before any Intune cmdlets instead of duplicating auth logic.
- Global folders (`BuildSpace`, `Scripts`, `Published`, `Temp`, etc.) are injected from [preferences.yaml](preferences.yaml); never hardcode paths because CI users relocate the repo.
- Interactive recipes are isolated under `Recipes/Interactive`; respect the `-NoInteractive` flag so unattended runs stay headless.
- This project uses a custom fork of [IntuneWin32App](https://github.com/jeffpearson1/IntuneWin32App) with a number of improvements including expanded Windows 11 version support, fixing "none" architecture handling and replacing errant uses of Write-Warning with Write-Error among other things.

## Configuration & Recipes
- Copy [Preferences.example.yaml](Preferences.example.yaml) to `preferences.yaml` and fill Tenant/Client secrets plus folder roots; defaults (scope tags, deployment groups, install behavior) are merged by `Set-ScriptVariables`.
- Recipes in [Recipes](Recipes) are plain YAML; the `id` must match the file name, and defaults can be overridden per recipe (see [Recipes/googlechrome.yaml](Recipes/googlechrome.yaml)).
- Placeholders `<filename>`, `<version>`, `<productcode>` are expanded by `Update-ScriptPlaceholders`; include them in install/uninstall/detection blocks instead of interpolating manually.
- Use [RecipeGroups.yaml](RecipeGroups.yaml) when wiring the `-Group` parameter so shared bundles (e.g., the Adobe set) stay declarative.
- `versionLock` accepts dotted patterns with `x` wildcards; `Compare-AppVersions` and `Test-VersionExcluded` enforce the rule before any download.

## Recipe Authoring Guidelines
- Start from the templates under [Templates](Templates) (e.g., [Templates/EXETemplate.yaml](Templates/EXETemplate.yaml) or `MSITemplate.yaml`) so you inherit the standard fields (`installExperience`, `restartBehavior`, `minOSVersion`) and only fill in what the vendor actually needs.
- Keep naming consistent: the recipe file name (e.g., `googlechrome.yaml`), the `id`, and the Intune display name prefix must align so `Get-SameAppAllVersions` can link releases; use all-lowercase for filenames and `id` but allow display names with spaces (`displayName: Google Chrome`).
- Populate download metadata in order of increasing complexity: use a static `url` when possible, fall back to `preDownloadScript` for lightweight metadata scrapes like [Recipes/googlechrome.yaml](Recipes/googlechrome.yaml) and [Recipes/7zip.yaml](Recipes/7zip.yaml), and reserve `downloadScript` for full custom flows (cookies, Selenium, authenticated sources). Always set `$fileName`, `$version`, and `$fileDetectionVersion` inside these scripts so packaging and detection stay in sync.
- Author install/uninstall commands with placeholders (`msiexec /i <filename> /qn /norestart`, `msiexec /x <productcode> /qn /norestart`) so Yardstick can inject the resolved artifact; add `postDownloadScript` only when you must modify the payload (unzipping, repacking with additional files).
- Choose the right detection strategy: `detectionType: msi` pulls ProductCode automatically like [Recipes/7zip.yaml](Recipes/7zip.yaml), while EXE/portable apps should supply `fileDetectionPath`, `fileDetectionName`, `fileDetectionMethod`, and `fileDetectionVersion` as shown in [Templates/EXETemplate.yaml](Templates/EXETemplate.yaml). Use `softwareName` when you rely on Add/Remove Programs queries.
- Capture lifecycle metadata per recipe: set `fileType` (`msi`, `EXE`, `zip`), `numVersionsToKeep`, `scopeTags`, and optional requirement overrides (`architecture`, `minOSVersion`) so `Move-AssignmentsAndDependencies` and retention pruning behave deterministically.
- Store icons in [Icons](Icons) and reference them via `iconFile`; never commit binaries under Recipes. If a vendor requires helper content (wrappers, transforms), keep those scripts under [Scripts](Scripts) or `BuildSpace` during execution so `Invoke-Cleanup` can purge them between runs.
- For MSI installers, prefer the fields in [Templates/MSITemplate.yaml](Templates/MSITemplate.yaml): keep `fileType: msi`, declare `detectionType: msi`, and let Yardstick call `Get-MsiProductCode` rather than hardcoding GUIDs. Recipes such as [Recipes/googlechrome.yaml](Recipes/googlechrome.yaml) and [Recipes/zoomrooms.yaml](Recipes/zoomrooms.yaml) show the pattern—populate `url` or `preDownloadScript` only, leave `<productcode>` placeholders in uninstall commands, and avoid extra detection metadata unless the vendor ships multiple MSI payloads.
- Vendors that ship chained MSI installers (e.g., [Recipes/qgis.yaml](Recipes/qgis.yaml)) often need custom `installScript` arguments. Still keep `<filename>` and `<productcode>` placeholders intact so `Update-ScriptPlaceholders` can rewrite them when Yardstick pins the download.
- The preDownloadScript is necessary for setting any variables used in the downloadScript that cannot be static. For example, the download script always requires a version, and since the version is always dynamic it must be set in the preDownloadScript. Note that `$version` **must** be set in `preDownloadScript` (not `downloadScript`) because version validation runs between the two scripts.
- All script blocks (`preDownloadScript`, `downloadScript`, `postDownloadScript`, `postRunScript`) execute with `-NoNewScope`, meaning they share the same scope and can read/write each other's variables.
- **Required variables** that must always be set by `preDownloadScript` (or statically in YAML):
  - `$version` – The application version string used for Intune display name and version comparison (e.g., `131.0.6778.109`, `24.08`). Must be set before version validation, which runs after `preDownloadScript` but before `downloadScript`.
  - `$fileName` – The exact name of the downloaded file (e.g., `ChromeSetup.msi`, `7z2408-x64.exe`). Can be set in either `preDownloadScript` or `downloadScript`, or statically in YAML.
- **Conditional variables** required based on recipe configuration:
  - `$url` – Required when there is no `downloadScript`; the direct download URI used by the default BITS transfer. Can be set statically in YAML or dynamically in `preDownloadScript`.
  - `$productCode` – Automatically extracted by `Get-MsiProductCode` after download for MSI files. Rarely needs to be set manually.
  - `$fileDetectionVersion` – The version string used in file-based and MSI detection rules. Defaults to `$version` if not set. Only set explicitly when the detection version format differs from the application version (e.g., needs 4-part padding via `[VersionPro]::new($version).ToString(4)`).
  - **File detection variables** (required when `detectionType: file`; can be set in YAML or scripts):
    - `$fileDetectionPath` – Full path where the installed file resides (e.g., `C:\Program Files\7-Zip`)
    - `$fileDetectionName` – The specific file name to detect (e.g., `7z.exe`)
    - Note: `fileDetectionMethod` and `fileDetectionOperator` are typically set in the YAML recipe itself, not in scripts
  - **Registry detection variables** (required when `detectionType: registry`; can be set in YAML or scripts):
    - `$registryDetectionKey` – Registry path to check (e.g., `HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{GUID}`)
    - `$registryDetectionValueName` – Name of the registry value to detect
    - `$registryDetectionValue` – Expected value for comparison (when using value-based detection)
    - Note: `registryDetectionOperator` and `registryDetectionMethod` are typically set in the YAML recipe
  - **Script detection variables** (required when `detectionType: script`):
    - `$detectionScript` – Set in YAML; the PowerShell script content that performs custom detection logic. Supports the `<version>` placeholder.

## Packaging & Detection
- Every run purges and recreates `BuildSpace/<id>/<version>`; scripts should write temporary assets there to ensure `Invoke-Cleanup` can delete them safely.
- Downloads default to BITS (`Start-BitsTransfer`); supply custom PowerShell via `downloadScript` when vendors need cookies or Selenium sessions (Adobe automation lives in [Modules/Custom](Modules/Custom)).
- Detection rules are centralized in `New-DetectionRule`; set `detectionType` plus supporting fields (`fileDetectionMethod`, `registryDetectionOperator`, etc.) in the recipe instead of crafting Intune objects inline.
- MSI packages automatically read ProductCode via `Get-MsiProductCode`; avoid hardcoding GUIDs in recipes.
- Requirement rules always come from `MinOSVersion` + `Architecture`; align recipe values with Intune-supported enums (e.g., `W10_20H2`).

## Intune Assignments & Cleanup
- `Move-AssignmentsAndDependencies` migrates group assignments, availability/deadline offsets, filters, and dependencies from older app versions to the new one; call it rather than re-implementing Graph moves.
- Naming convention is enforced (`DisplayName`, `DisplayName (N-1)`, etc.) via `Get-SameAppAllVersions`; any custom rename logic must preserve this scheme or repairs will run every execution.
- `defaultDeploymentGroups` ensures new builds auto-assign to baseline groups; the script only adds missing assignments, so safe to run repeatedly.
- Respect `-Force`, `-Repair`, and `-NoDelete` flags when extending the workflow; they toggle version comparisons, naming cleanup, and retention pruning inside the main loop.

## Extensibility & Secrets
- Custom helpers dropped in [Modules/Custom](Modules/Custom) are auto-imported; use this for vendor-specific download automation (see [Modules/Custom/AdobeDownloader.psm1](Modules/Custom/AdobeDownloader.psm1) for Selenium patterns and credential usage via `Invoke-SSOAutoSignIn`).
- `Get-Secrets` expects YAML vaults under [Secrets](Secrets); return hashtables keyed by whatever your module needs rather than re-reading preference files.
- Version sorting uses the custom [Modules/VersionPro.psm1](Modules/VersionPro.psm1) type; cast versions to `[VersionPro]` whenever you need consistent ordering of semantic-ish strings.

## Reporting & Tooling
- Email summaries are generated in `Send-YardstickEmailReport` (Outlook COM) or `Send-YardstickEmailReportMailKit` when `emailDeliveryMethod` is set; populate the Email section of preferences before enabling `emailNotificationEnabled`.
- Use [Test-EmailNotification.ps1](Test-EmailNotification.ps1) with `-TestReport` to validate HTML output and Outlook connectivity without pushing new apps.
- Icons are never committed; place 512x512 PNG/JPG files in [Icons](Icons) with the exact `iconFile` name referenced in recipes to avoid runtime failures.
- Selenium-driven recipes assume Firefox plus matching `geckodriver`; keep the driver binaries in [Tools](Tools) (curl.exe also lives there for Adobe scraping).
- Always finish runs with `.\Yardstick.ps1 -All -NoInteractive -NoEmail` locally before committing recipe changes so detection, packaging, and assignment migrations are exercised end-to-end.
