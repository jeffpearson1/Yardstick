# Writing Yardstick Recipes

This guide provides comprehensive documentation for creating new application recipes in Yardstick.

## Quick Start

1. Copy the appropriate template from the `Templates/` folder:
   - `MSITemplate.yaml` for MSI installers
   - `EXETemplate.yaml` for EXE installers or portable applications
2. Rename the file to match your application name (e.g., `googlechrome.yaml`)
3. Fill in the required fields based on the sections below
4. Test your recipe with `.\Yardstick.ps1 -ApplicationId YourAppName -NoEmail`

## Recipe File Naming

The recipe filename (without `.yaml`) becomes the application ID. The filename must:
- Match the `id` field in the YAML file
- Use all-lowercase convention (e.g., `googlechrome`, `7zip`)
- Be unique across all recipes

## Core Required Fields

The following fields are **always required** by the recipe schema validator (`Test-RecipeSchema`):

- `id` - Must match the YAML filename without extension
- `displayName` - Human-readable name shown in Intune
- `detectionType` - One of: `msi`, `file`, `registry`, `script`
- `iconFile` - Icon filename in the `Icons/` folder (PNG or JPG, max 512x512)
- `description` - Brief description of the application
- `publisher` - Software vendor or publisher name
- One of `installScript` or `powerShellInstallScript`
- One of `uninstallScript` or `powerShellUninstallScript`

> **Note:** Required fields can be satisfied either in the YAML itself or by setting the corresponding variable (e.g., `$registryDetectionKey = ...`) in a script block. The schema validator checks both.

### Functionally Required (Not Schema-Validated)

These fields are not checked by the schema validator but are required at runtime:

- **`$version`** - Must be set in `preDownloadScript` or the YAML `version` field. Validated immediately after `preDownloadScript` runs, so it **cannot** be set in `downloadScript`.
- **`$fileName`** - Must be set in YAML, `preDownloadScript`, or `downloadScript`. Required for `.intunewin` packaging and placeholder replacement.
- **`$url`** - Must be set in YAML, `preDownloadScript`, or `downloadScript` unless a `downloadScript` handles the download entirely.

## Script Blocks

All script blocks execute with `-NoNewScope`, meaning they run directly in the caller's scope. This has two important implications:

1. **All scripts share the same scope.** Variables set in `preDownloadScript` are visible to `downloadScript`, `postDownloadScript`, and `postRunScript`.
2. **Any `$Script:` variable can be set or overridden.** Writing `$version = "1.2.3"` in a script block writes directly to `$Script:Version` (PowerShell is case-insensitive).

Scripts execute in this order:

1. `preDownloadScript` - Before download and version validation
2. `downloadScript` - Replaces default BITS download (working directory set to `$BuildSpace\$id\$version`)
3. `postDownloadScript` - After download completes (working directory set to `$BuildSpace\$id\$version`)
4. `postRunScript` - After Intune upload (runs outside the try/catch block)

### Environment Variables Available to Scripts

These variables are set once at startup from `preferences.yaml` and are available in all script blocks:

| Variable | Description |
|---|---|
| `$BuildSpace` | Staging directory for downloads (files go to `$BuildSpace\$id\$version\`) |
| `$Temp` | Temp directory |
| `$Scripts` | Scripts directory |
| `$Published` | Output directory for .intunewin files |
| `$Recipes` | Recipes directory |
| `$Icons` | Icons directory |
| `$Tools` | Tools directory |
| `$Secrets` | Secrets directory |

### `preDownloadScript`

Executes before any download processing and before version validation. Use this to:
- Scrape version numbers from vendor websites
- Construct dynamic download URLs
- Set version-dependent variables

**Must set:**
- `$version` - The application version (validated immediately after this script runs)

**Commonly set:**
- `$url` - Download URL when not static in YAML
- `$fileName` - When filename includes version or is dynamic
- `$fileDetectionVersion` - Version string for file or MSI detection rules (defaults to `$version` if not set). May need to be padded to 4 parts: `$fileDetectionVersion = [VersionPro]::new($version).ToString(4)`

**Example (7-Zip):**
```yaml
preDownloadScript: |
  $html = $(Invoke-WebRequest "https://www.7-zip.org/download.html").content
  $html -match "a\/7z[0-9]{1,6}-x64\.msi" | Out-Null
  $url = "https://www.7-zip.org/$($matches[0])"
  $fileName = $($matches[0] -split "/")[1]
  $html -match "[0-9]{2}\.[0-9]{2}" | Out-Null
  $version = $matches[0]
  $fileDetectionVersion = $version
```

### `downloadScript`

Custom download logic. Yardstick uses BITS transfer by default, but this script replaces it entirely. The working directory is set to `$BuildSpace\$id\$version` before execution.

Use for:
- Custom authentication
- Cookie handling
- Multi-file downloads
- Selenium-based automation

**Must set (if not already set in `preDownloadScript`):**
- `$fileName` - Name of the downloaded file

**Important:** `$version` **cannot** be set here -- version validation happens between `preDownloadScript` and `downloadScript`. Always set `$version` in `preDownloadScript`.

**Must do:**
- Perform the actual file download (the default BITS transfer is skipped when this script is present)

### `postDownloadScript`

Executes after download completes. The working directory is set to `$BuildSpace\$id\$version` before execution.

Use for:
- Extracting archives
- Modifying installer files
- Adding additional files to the package
- Post-processing downloaded content

**Example:**
```yaml
postDownloadScript: |
  Expand-Archive -Path $fileName -DestinationPath ".\extracted"
  Copy-Item "$Scripts\config.xml" -Destination "."
```

### `postRunScript`

Executes after the Intune upload completes. Runs outside the main try/catch block, so errors here do not prevent processing the next application.

Use for:
- Cleanup operations
- Logging
- Notifications
- Any post-deployment tasks

## Installation and Uninstallation

### MSI Installers

**`installScript`** (string) - MSI install command.
- Standard: `installScript: msiexec /i <filename> /qn /norestart`
- The `<filename>` placeholder is automatically replaced

**`uninstallScript`** (string) - MSI uninstall command.
- Standard: `uninstallScript: msiexec /x <productcode> /qn /norestart`
- The `<productcode>` placeholder is automatically populated from the MSI

### EXE Installers

**`installScript`** (string) - EXE install command with silent switches.
- Example: `installScript: <filename> /S`
- Example: `installScript: <filename> /VERYSILENT /NORESTART`

**`uninstallScript`** (string) - EXE uninstall command.
- Example: `uninstallScript: '"C:\Program Files\App\uninstall.exe" /S'`
- Use proper escaping for paths with spaces

### PowerShell Install Scripts

**`powerShellInstallScript`** (string) - PowerShell script content for complex installations.
- Automatically creates `install.ps1` in the package
- Auto-generates `installScript` as: `powershell.exe -noprofile -executionpolicy bypass -file .\install.ps1`
- Only auto-generates if `installScript` is not also set

**`powerShellUninstallScript`** (string) - PowerShell script content for complex uninstallations.
- Automatically creates `uninstall.ps1` in the package
- Auto-generates `uninstallScript` as: `powershell.exe -noprofile -executionpolicy bypass -file .\uninstall.ps1`
- Only auto-generates if `uninstallScript` is not also set

**Example (QGIS):**
```yaml
powershellInstallScript: |
  Invoke-WebRequest -Uri "https://github.com/google/fonts/raw/main/OpenSans.ttf" -UseBasicParsing
  if (Test-Path -Path "$($env:ProgramData)\Microsoft\Windows\Start Menu\Programs\QGIS*") {
    Remove-Item -Path "$($env:ProgramData)\Microsoft\Windows\Start Menu\Programs\QGIS*" -Recurse -Force
  }
  Start-Process -FilePath "msiexec" -ArgumentList "/i `"<filename>`" INSTALLDIR=`"$env:ProgramFiles\QGIS`" /qn /norestart" -Wait
```

## Detection Rules

Yardstick supports four detection types: `msi`, `file`, `registry`, and `script`.

### MSI Detection

**`detectionType: msi`**

The simplest detection method. Automatically extracts and uses the MSI ProductCode via `Get-MsiProductCode` after download.

**Required fields:**
- `detectionType: msi`

**Optional fields:**
- `fileDetectionVersion` (string) - Overrides the MSI product version used in detection. Defaults to `$version` if not set.

**Example:**
```yaml
detectionType: msi
```

### File Detection

**`detectionType: file`**

Detects based on file existence, version, date, or size.

**Required fields:**
- `detectionType: file`
- `fileDetectionPath` (string) - Full folder path where the file exists
- `fileDetectionName` (string) - Filename to detect
- `fileDetectionMethod` (string) - Detection method: `version`, `exists`, `modified`, `created`, `size`

**Version detection fields:**
- `fileDetectionVersion` (string) - Expected version. Defaults to `$version` if not set. Set in `preDownloadScript` when version padding is needed.
- `fileDetectionOperator` (string) - Comparison operator: `equal`, `notEqual`, `greaterThan`, `greaterThanOrEqual`, `lessThan`, `lessThanOrEqual`

**Date detection fields:**
- `fileDetectionDateTime` (datetime) - Date/time to compare against
- `fileDetectionOperator` (string) - Same operators as version

**Size detection fields:**
- `fileDetectionValue` (integer) - File size in **MB**
- `fileDetectionOperator` (string) - Same operators as version

**Example (Notepad++):**
```yaml
detectionType: file
fileDetectionPath: 'C:\Program Files\Notepad++'
fileDetectionName: notepad++.exe
fileDetectionMethod: version
fileDetectionVersion:  # Set in preDownloadScript; defaults to $version if omitted
fileDetectionOperator: equal
```

### Registry Detection

**`detectionType: registry`**

Detects based on registry key/value existence or value content.

**Required fields:**
- `detectionType: registry`
- `registryDetectionMethod` (string) - Detection method: `exists`, `notExists`, `string`, `integer`, `version`
- `registryDetectionKey` (string) - Full registry path including hive

**Value-based detection fields:**
- `registryDetectionValueName` (string) - Name of the registry value
- `registryDetectionValue` (string/int) - Expected value
- `registryDetectionOperator` (string) - Comparison operator (same as file detection)

> **Note:** Registry detection fields can be set dynamically in scripts. For example, `$registryDetectionKey` can be set in `preDownloadScript` when the key path includes a version number. See `djv.yaml` for this pattern.

**Example (Firefox):**
```yaml
detectionType: registry
registryDetectionMethod: exists
registryDetectionKey: 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Mozilla Firefox <version> (x64 en-US)'
```

**Note:** The `<version>` placeholder in registry keys is automatically replaced by `Update-ScriptPlaceholders`.

### Script Detection

**`detectionType: script`**

Custom PowerShell detection script for complex scenarios.

**Required fields:**
- `detectionType: script`
- `detectionScript` (string) - PowerShell script content

**Optional fields:**
- `detectionScriptFileExtension` (string) - Script file extension (default from preferences)
- `detectionScriptRunAs32Bit` (boolean) - Run in 32-bit context (default from preferences)
- `detectionScriptEnforceSignatureCheck` (boolean) - Require signed script (default from preferences)

**Example (Python 3):**
```yaml
detectionType: script
detectionScript: |
  $version = "<version>"
  $dirVersion = (($version.Split('.')[0,1] -join '.').Replace('.',''))
  $pythonExe = Join-Path $env:ProgramFiles ("Python" + $dirVersion + "\python.exe")
  $installedVersion = $null
  if (Test-Path $pythonExe) {
      $installedVersion = (Get-Item -Path $pythonExe -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
  }
  if ($installedVersion -eq "$version") {
      Write-Output "Detected"
      exit 0
  }
  Write-Output "Not Detected"
  exit 1
```

> **Note:** The `<version>` placeholder in `detectionScript` is replaced by `Update-ScriptPlaceholders`, so you can use it to inject the current version at packaging time.

## Installation Behavior Settings

These settings control how Intune handles installation. All default to values in `preferences.yaml` when omitted.

**`installExperience`** (string) - Installation context.
- Values: `system`, `user`
- Example: `installExperience: system`

**`restartBehavior`** (string) - Restart handling.
- Values: `allow`, `suppress`, `force`, `basedOnReturnCode`
- Example: `restartBehavior: suppress`

**`maximumInstallationTimeInMinutes`** (integer) - Installation timeout.
- Example: `maximumInstallationTimeInMinutes: 60`

**`allowUserUninstall`** (boolean) - Allow users to uninstall.
- Example: `allowUserUninstall: false`

## System Requirements

**`minOSVersion`** (string) - Minimum Windows version required.
- Values: `W10_1607`, `W10_1703`, `W10_1709`, `W10_1803`, `W10_1809`, `W10_1903`, `W10_1909`, `W10_2004`, `W10_20H2`, `W10_21H1`, `W10_21H2`, `W10_22H2`, `W11_21H2`, `W11_22H2`, `W11_23H2`, `W11_24H2`
- Default: From `preferences.yaml`
- Example: `minOSVersion: W10_20H2`

**`architecture`** (string) - Target architecture.
- Values: `x64`, `x86`, `arm64`, `none`
- Default: From `preferences.yaml`
- Example: `architecture: x64`

**`is32BitApp`** (boolean) - Indicates if application is 32-bit.
- Default: From `preferences.yaml`
- Example: `is32BitApp: false`

## Deployment Configuration

**`scopeTags`** (array) - Intune scope tags for RBAC.
- Default: From `preferences.yaml`
- Example:
  ```yaml
  scopeTags:
    - My Scope Tag
    - My Scope Tag 2
  ```

**`owner`** (string) - Application owner.
- Default: From `preferences.yaml`
- Example: `owner: IT Department`

**`defaultDeploymentGroups`** (array) - Azure AD groups for automatic assignment.
- Default: From `preferences.yaml`
- Example:
  ```yaml
  defaultDeploymentGroups:
    - All Users
    - IT Pilot Group
  ```

**`availableGroups`** (array) - Groups for "Available" assignment intent.
- Default: From `preferences.yaml`

**`requiredGroups`** (array) - Groups for "Required" assignment intent.
- Default: From `preferences.yaml`

**`availableDateOffset`** (integer) - Days from now when app becomes available.
- Default: From `preferences.yaml`
- Example: `availableDateOffset: 0`

**`deadlineDateOffset`** (integer) - Days from now for required installation deadline.
- Default: From `preferences.yaml`
- Example: `deadlineDateOffset: 7`

## Version Management

**`numVersionsToKeep`** (integer) - Number of app versions to retain in Intune.
- Default: From `preferences.yaml`
- Example: `numVersionsToKeep: 2`
- Note: When `versionLock` is set, this is forced to 1

**`versionLock`** (string) - Lock to specific version pattern using wildcards.
- Pattern: Use `x` for wildcard positions
- Example: `versionLock: 3.11.x` (only deploy Python 3.11.x versions)
- Example: `versionLock: 2024.x.x` (only deploy 2024.x.x versions)

**`displayVersion`** (string) - Override version shown in Intune (optional).
- Useful when file version differs from marketing version

## Recipe Inheritance

Recipes support single-level inheritance via the `base` field. When a recipe contains a `base` field, the referenced base recipe is loaded and the child recipe's fields are overlaid on top.

**`base`** (string) - The `id` of another recipe to inherit from.
- Child fields override base fields (shallow merge)
- Chained inheritance is not supported (a base recipe cannot have its own `base` field)
- The `base` field itself is excluded from the merged result

**Example:**
```yaml
base: photoshop
id: photoshop_sdl
displayName: Adobe Photoshop (SDL)
scopeTags:
  - Different Scope Tag
```

This inherits all fields from `photoshop.yaml` and overrides only `id`, `displayName`, and `scopeTags`.

## Dependency Management

**`allowDependentLinkUpdates`** (boolean) - Allow updating dependency links when creating new versions.
- Default: From `preferences.yaml`
- Example: `allowDependentLinkUpdates: true`

**`dependentLinkUpdateEnabled`** (boolean) - Enable dependency link update feature.
- Default: From `preferences.yaml`

**`dependentLinkUpdateRetryCount`** (integer) - Number of retry attempts for dependency updates.
- Default: From `preferences.yaml` (typically 3)

**`dependentLinkUpdateRetryDelaySeconds`** (integer) - Delay between retry attempts.
- Default: From `preferences.yaml` (typically 5)

**`dependentLinkUpdateTimeoutSeconds`** (integer) - Timeout for dependency updates.
- Default: From `preferences.yaml` (typically 60)

**`dependentApplicationBlacklist`** (array) - Application IDs to exclude from dependency updates.
- Default: From `preferences.yaml`
- Example:
  ```yaml
  dependentApplicationBlacklist:
    - LegacyApp
    - TestApp
  ```

## Placeholders

`Update-ScriptPlaceholders` automatically replaces placeholders in these fields: `installScript`, `powerShellInstallScript`, `uninstallScript`, `powerShellUninstallScript`, `detectionScript`, and `registryDetectionKey`.

The three placeholders are:

| Placeholder | Replaced With | Notes |
|---|---|---|
| `<filename>` | `$fileName` | The downloaded file name |
| `<version>` | `$version` | The application version |
| `<productcode>` | MSI ProductCode | Only populated for MSI packages (extracted by `Get-MsiProductCode` after download) |

Placeholder replacement runs three times during processing: before download, after download, and after MSI product code extraction. If any placeholders remain unreplaced after all three passes, the recipe fails with an error.

**Example:**
```yaml
installScript: msiexec /i <filename> /qn /norestart
uninstallScript: msiexec /x <productcode> /qn /norestart
```

Becomes:
```powershell
msiexec /i ChromeSetup.msi /qn /norestart
msiexec /x {8A69D345-D564-463C-AFF1-A69D9E530F96} /qn /norestart
```

## Complete Examples

### MSI Installer Example (Google Chrome)

```yaml
url: https://dl.google.com/chrome/install/googlechromestandaloneenterprise64.msi
urlredirects: false
preDownloadScript: |
  $VURL = "https://versionhistory.googleapis.com/v1/chrome/platforms/win64/channels/stable/versions/all/releases?filter=endtime=none&order_by=fraction%20desc"
  $Content = Invoke-WebRequest $VURL
  $JSON = $Content.content | ConvertFrom-JSON
  $Version = $JSON.releases[0].version
  $fileDetectionVersion = $Version
installScript: msiexec /i <filename> /qn /norestart
uninstallScript: msiexec /x <productcode> /qn /norestart
fileName: googlechromestandaloneenterprise64.msi
iconFile: GoogleChrome.png
id: GoogleChrome
displayName: Google Chrome
publisher: Google
description: A web browser by Google
detectionType: msi
numVersionsToKeep: 2
scopeTags:
  - My Scope Tag
  - My Scope Tag 2
```

### EXE Installer with File Detection Example (Notepad++)

```yaml
url:
urlredirects: false
preDownloadScript: |
  $html = Invoke-WebRequest "https://notepad-plus-plus.org/downloads/"
  $s = $($html.links | Select-String "Current Version") -replace ".{0,}Current Version "
  $version = ($s.Substring(0, $s.IndexOf(";")) -split "<")[0]
  $url = "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v$version/npp.$version.Installer.x64.exe"
  $fileName = "npp.$version.Installer.x64.exe"
  $fileDetectionVersion = $version
installScript: <filename> /S
uninstallScript: '"C:\Program Files\Notepad++\uninstall.exe" /S'
iconFile: npp.png
id: notepadplusplus
displayName: Notepad++
publisher: Don Ho
description: 'A free source code editor and Notepad replacement'
detectionType: file
fileDetectionPath: 'C:\Program Files\Notepad++'
fileDetectionName: notepad++.exe
fileDetectionMethod: version
fileDetectionOperator: equal
scopeTags:
  - My Scope Tag
  - My Scope Tag 2
```

### Registry Detection Example (Firefox)

```yaml
url: https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win64&lang=en-US
urlredirects: true
preDownloadScript: |
  $filename = $URL.split("/")[-1].replace("%20", " ")
  $version = ($filename.split(" ")[-1] -Split "\.msi")[0]
installScript: msiexec /i "<filename>" /qn /norestart
uninstallScript: |
  "%ProgramFiles%\Mozilla Firefox\uninstall\helper.exe" -ms
iconFile: firefox.png
id: firefox
displayName: Mozilla Firefox
publisher: The Mozilla Foundation
description: Mozilla Firefox is a free and open-source web browser
detectionType: registry
registryDetectionMethod: exists
registryDetectionKey: 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Mozilla Firefox <version> (x64 en-US)'
numVersionsToKeep: 2
scopeTags:
  - My Scope Tag
  - My Scope Tag 2
```

### Registry Detection with Dynamic Keys Example (DJV)

```yaml
url: https://github.com/grizzlypeak3d/DJV
urlredirects: false
preDownloadScript: |
  using module ".\Modules\Custom\GithubDownloader.psm1"
  $downloadFileRegex = "djv-.*amd64\.exe$"
  $versionRegex = "(\.{0,1}[0-9]{1,2}){3}"
  $Script:downloader = [GithubDownloader]::new($URL, $downloadFileRegex, $versionRegex)
  $downloader.Update()
  $Version = $downloader.LatestVersion
  $URL = $downloader.URL
  $filename = $downloader.filename
  $registryDetectionKey = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\DJV $Version"
  $registryDetectionValue = $Version
installScript: <filename> /S
uninstallScript: '"%ProgramFiles%\DJV <version>\Uninstall.exe" /S'
iconFile: djv.png
id: djv
displayName: DJV
publisher: Grizzly Peak Software
description: DJV is an open source application for media playback and review.
detectionType: registry
registryDetectionValueName: DisplayVersion
registryDetectionOperator: greaterThanOrEqual
registryDetectionMethod: version
```

### Script Detection Example (Python 3)

```yaml
url:
urlredirects: false
preDownloadScript: |
  $html = $(Invoke-WebRequest "https://www.python.org/downloads/windows/").Content
  $html -match "Latest Python 3 Release - Python 3\.[0-9]{1,3}\.[0-9]{1,3}" | Out-Null
  $version = $($($matches[0] -split "-")[1]).Trim(" Python")
  $url = "https://www.python.org/ftp/python/$version/python-$version-amd64.exe"
  $fileName = "python-$version-amd64.exe"
installScript: '<filename> /quiet InstallAllUsers=1 PrependPath=1'
uninstallScript: '<filename> /quiet /uninstall'
iconFile: python3.png
id: python3
displayName: Python 3
publisher: Python Software Foundation
description: Python is a programming language that lets you work quickly and integrate systems effectively.
detectionType: script
detectionScript: |
  $version = "<version>"
  $dirVersion = (($version.Split('.')[0,1] -join '.').Replace('.',''))
  $pythonExe = Join-Path $env:ProgramFiles ("Python" + $dirVersion + "\python.exe")
  $installedVersion = $null
  if (Test-Path $pythonExe) {
      $installedVersion = (Get-Item -Path $pythonExe -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
  }
  if ($installedVersion -eq "$version") {
      Write-Output "Detected"
      exit 0
  }
  Write-Output "Not Detected"
  exit 1
```

## Testing Your Recipe

1. **Test the download and version detection:**
   ```powershell
   .\Yardstick.ps1 -ApplicationId YourAppName -NoEmail
   ```

2. **Test with force flag to recreate existing version:**
   ```powershell
   .\Yardstick.ps1 -ApplicationId YourAppName -Force -NoEmail
   ```

3. **Test repair functionality:**
   ```powershell
   .\Yardstick.ps1 -ApplicationId YourAppName -Repair -NoEmail
   ```

4. **Verify detection rules work correctly in Intune after deployment**

## Common Pitfalls

1. **Setting `$version` in `downloadScript`** - Version validation runs *before* `downloadScript`. Always set `$version` in `preDownloadScript` or the YAML `version` field.
2. **Incorrect file detection version format** - Ensure it matches the actual file properties. Use `[VersionPro]::new($version).ToString(4)` to pad to 4-part format when needed.
3. **Missing icon files** - Icons must exist in `Icons/` folder with the exact name referenced in `iconFile`.
4. **Hardcoded paths** - Use `<filename>`, `<version>`, and `<productcode>` placeholders instead of hardcoding values in install/uninstall scripts.
5. **Case sensitivity in YAML keys** - YAML keys are case-sensitive. Use the exact casing from this guide (e.g., `urlredirects` not `urlRedirects`, `fileDetectionMethod` not `filedetectionmethod`). The schema validator warns about casing mismatches.
6. **Assuming `$fileDetectionVersion` is required** - It defaults to `$version` when not set. Only set it explicitly when the file version format differs from the application version.
7. **Wrong detection type** - MSI packages should use `detectionType: msi` unless you need custom detection logic. MSI detection automatically extracts the ProductCode.
8. **Unreplaced placeholders** - If any `<filename>`, `<version>`, or `<productcode>` placeholders remain in scripts after processing, the recipe will fail. Ensure the corresponding variables are set.

## Advanced Topics

### Recipe Inheritance

See the [Recipe Inheritance](#recipe-inheritance) section above. This pattern is commonly used for Adobe products where an SDL (Software Distribution License) variant inherits from the base recipe and overrides only `id`, `displayName`, and scope-related fields.

### Version Normalization

Some vendors use inconsistent version formats. Normalize in `preDownloadScript`:

```yaml
preDownloadScript: |
  $version = "24.08"
  # Convert to 4-part version for file detection
  $fileDetectionVersion = [VersionPro]::new($version).ToString(4)
```

### Custom Download Authentication

For applications requiring authentication, use `downloadScript`:

```yaml
downloadScript: |
  $headers = @{ Authorization = "Bearer $token" }
  Invoke-WebRequest -Uri $url -Headers $headers -OutFile "$BuildSpace\$id\$version\$fileName"
```

### Selenium-Based Downloads

For complex download scenarios (JavaScript-heavy pages, multi-step auth), see `Modules/Custom/AdobeDownloader.psm1` for Selenium patterns.

## Getting Help

- Review existing recipes in `Recipes/` folder for examples
- Check templates in `Templates/` folder
- Review Copilot instructions in `.github/copilot-instructions.md`
